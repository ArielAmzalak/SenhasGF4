
# event_utils.py — utilitários para o "Distribuidor de Senhas" (Streamlit + Google Sheets + PDF)
from __future__ import annotations

import io
import os
import re
import json
import unicodedata
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from datetime import datetime
from zoneinfo import ZoneInfo

import qrcode
from fpdf import FPDF
from barcode import Code128
from barcode.writer import ImageWriter

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.service_account import Credentials as SACredentials
from google.oauth2.credentials import Credentials as UserCredentials
from google.auth.transport.requests import Request

try:
    import streamlit as st
except ModuleNotFoundError:
    st = None

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
DEFAULT_TIMEZONE = os.getenv("APP_TZ", "America/Manaus")

HEADERS = [
    "Senha",
    "Nome",
    "Telefone",
    "Rede Social",
    "E-mail",
    "Bairro",
    "Data e Hora de Registro",
    "Data e Hora de Atendimento",
]


def _column_letter(idx: int) -> str:
    """Converte um índice baseado em zero para a letra de coluna do Sheets/Excel."""

    idx += 1
    letters = []
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters.append(chr(ord("A") + rem))
    return "".join(reversed(letters))


def _header_range(title: str) -> str:
    last_col = _column_letter(len(HEADERS) - 1)
    return f"{title}!A1:{last_col}1"

# Aba com as áreas/setores
NOMES_SHEET = os.getenv("NOMES_SHEET", "Nomes")

# Aba com a lista de bairros
BAIRROS_SHEET = os.getenv("BAIRROS_SHEET", "Bairro")

# ✅ Pedido do usuário: Spreadsheet ID definido **no código** (não em secrets)
HARDCODED_SPREADSHEET_ID = "1eEvF5c8rTXwWKqgmyCMXU5OPJKqBk5XPt4Yry5B4x5c"

DEFAULT_LOGO_PATH = Path(__file__).resolve().parent / "assets" / "logo.png"
PDF_LOGO_PATH = os.getenv("PDF_LOGO_PATH")


def _normalize(s: str) -> str:
    s = s or ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.strip().lower()


def format_phone_number(telefone: str) -> str:
    """Normaliza números para o padrão local `(92) 98123-1234`.

    Levanta ``ValueError`` quando o telefone não possui 11 dígitos (incluindo DDD).
    """

    telefone = (telefone or "").strip()
    if not telefone:
        raise ValueError("Telefone é obrigatório.")

    digits = re.sub(r"\D", "", telefone)
    if not digits:
        raise ValueError("Telefone deve conter apenas números válidos.")

    # Remove código do país (55) se presente, mantendo 11 dígitos finais
    if digits.startswith("55") and len(digits) > 11:
        digits = digits[2:]

    if len(digits) != 11:
        raise ValueError("Telefone deve conter 11 dígitos (incluindo DDD).")

    numero_local = digits[-9:]
    return f"(92) {numero_local[:5]}-{numero_local[5:]}"


def format_name_upper(nome: str) -> str:
    """Garante nome em caixa alta, preservando espaços externos."""

    return (nome or "").strip().upper()


def _truthy(v: Any) -> bool:
    if isinstance(v, bool):
        return v
    s = _normalize(str(v))
    return s in {"sim", "s", "true", "1", "y", "yes", "ativo", "ativa", "on", "ok"}


def _parse_positive_int(value: Any) -> Optional[int]:
    """Converte valores variados para inteiro positivo, quando possível."""

    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        try:
            num = int(value)
        except (TypeError, ValueError):
            return None
    else:
        s = re.sub(r"\D", "", str(value))
        if not s:
            return None
        try:
            num = int(s)
        except ValueError:
            return None
    return num if num > 0 else None


def _authorize_google_sheets():
    # Prefer service account (GOOGLE_SERVICE_ACCOUNT_JSON) quando rodar na nuvem
    sa_json = None
    if st:
        sa_json = st.secrets.get("GOOGLE_SERVICE_ACCOUNT_JSON", None)
    if not sa_json:
        sa_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")

    if sa_json:
        try:
            info = json.loads(sa_json)
            creds = SACredentials.from_service_account_info(info, scopes=SCOPES)
            return creds
        except Exception as exc:
            raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON inválido.") from exc

    # Fallback: OAuth de usuário (GOOGLE_CLIENT_SECRET) — compatível com apps antigos
    client_json = None
    if st:
        client_json = st.secrets.get("GOOGLE_CLIENT_SECRET", None)
    if not client_json:
        client_json = os.getenv("GOOGLE_CLIENT_SECRET")
    if not client_json:
        raise RuntimeError("Credenciais Google ausentes. Defina GOOGLE_SERVICE_ACCOUNT_JSON (preferível) ou GOOGLE_CLIENT_SECRET.")

    token_path = "token.json"
    creds = None
    if os.path.exists(token_path):
        creds = UserCredentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            from google_auth_oauthlib.flow import InstalledAppFlow
            conf = json.loads(client_json)
            flow = InstalledAppFlow.from_client_config(conf, SCOPES)
            # Em ambiente headless, utiliza run_console()
            creds = flow.run_console()
        with open(token_path, "w", encoding="utf-8") as fp:
            fp.write(creds.to_json())
    return creds


def _sheets_service():
    return build("sheets", "v4", credentials=_authorize_google_sheets(), cache_discovery=False)


def _get_spreadsheet_id() -> str:
    # ✅ Preferir ID definido no código (não em secrets) — pedido do usuário
    sid = (HARDCODED_SPREADSHEET_ID or "").strip()
    if sid:
        return sid
    # Fallback para compatibilidade (caso remova o hardcoded)
    sid = None
    if st:
        sid = st.secrets.get("SPREADSHEET_ID")
    if not sid:
        sid = os.getenv("SPREADSHEET_ID", "")
    if not sid:
        raise RuntimeError("SPREADSHEET_ID não configurado (defina no código, secrets ou variável de ambiente).")
    return sid


def _find_col_indexes(header_row: List[str], candidates: List[str]) -> Optional[int]:
    norm = [_normalize(h) for h in header_row]
    for want in candidates:
        want_n = _normalize(want)
        for idx, col in enumerate(norm):
            if col == want_n:
                return idx
    return None


def _get_sheet_metadata(service, spreadsheet_id: str) -> Dict[str, Any]:
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return meta


def _sheet_exists(meta: Dict[str, Any], title: str) -> Tuple[bool, Optional[int]]:
    for s in meta.get("sheets", []):
        props = s.get("properties", {})
        if props.get("title") == title:
            return True, int(props.get("sheetId"))
    return False, None


def ensure_area_sheet(service, spreadsheet_id: str, title: str) -> None:
    """Garante que a aba da área existe e tem os cabeçalhos esperados."""
    meta = _get_sheet_metadata(service, spreadsheet_id)
    exists, _ = _sheet_exists(meta, title)
    if not exists:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": title}}}]},
        ).execute()
        # escreve cabeçalho
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=_header_range(title),
            valueInputOption="RAW",
            body={"values": [HEADERS]},
        ).execute()
        return
    # se já existe, valida cabeçalho (não falha se estiver diferente; apenas atualiza se vazio)
    rng = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{title}!1:1",
    ).execute()
    row1 = rng.get("values", [[]])
    if not row1 or not row1[0]:
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=_header_range(title),
            valueInputOption="RAW",
            body={"values": [HEADERS]},
        ).execute()


def read_active_areas(service, spreadsheet_id: str) -> List[Dict[str, Any]]:
    """
    Lê a aba NOMES (por padrão 'Nomes') e retorna apenas as áreas ativas.
    Campos aceitos (case/acento-insensitive):
      - Área (ou Area, Setor, Mesa)
      - Aba (ou Sheet, AbaDestino, Destino) — se ausente, usa o mesmo texto da Área
      - Ativa (ou Status) — valores: Sim/Nao, True/False, 1/0, Ativo/Inativo
    """
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{NOMES_SHEET}!A:Z",
        ).execute()
    except HttpError as exc:
        raise RuntimeError(f"Erro ao ler a aba '{NOMES_SHEET}': {exc}") from exc

    rows = result.get("values", [])
    if not rows:
        return []

    header = rows[0]
    area_idx = _find_col_indexes(header, ["Área", "Area", "Setor", "Mesa", "Área/Setor"])
    aba_idx = _find_col_indexes(header, ["Aba", "Sheet", "AbaDestino", "Aba Destino", "Destino", "Guia", "Tab"])
    ativa_idx = _find_col_indexes(header, ["Ativa", "Ativo", "Status", "Habilitada", "Disponível"])
    max_idx = _find_col_indexes(header, [
        "Quantidade máxima de senhas",
        "Qtd máxima",
        "Qtd Senhas",
        "Limite",
    ])

    if area_idx is None:
        raise RuntimeError("Coluna 'Área' (ou equivalente) não encontrada na aba 'Nomes'.")

    areas: List[Dict[str, Any]] = []
    for row in rows[1:]:
        area = (row[area_idx] if area_idx < len(row) else "").strip()
        if not area:
            continue
        sheet_title = (row[aba_idx] if (aba_idx is not None and aba_idx < len(row)) else area).strip() or area
        ativa_val = (row[ativa_idx] if (ativa_idx is not None and ativa_idx < len(row)) else "Sim")
        max_val = row[max_idx] if (max_idx is not None and max_idx < len(row)) else None
        ativa = _truthy(ativa_val)
        if ativa:
            areas.append({
                "area": area,
                "sheet": sheet_title,
                "ativa": True,
                "max_senhas": _parse_positive_int(max_val),
            })
    return areas


def read_neighborhoods(service, spreadsheet_id: str) -> List[str]:
    """Lê a aba de bairros e devolve uma lista com os nomes válidos."""
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{BAIRROS_SHEET}!A:A",
        ).execute()
    except HttpError as exc:
        raise RuntimeError(f"Erro ao ler a aba '{BAIRROS_SHEET}': {exc}") from exc

    rows = result.get("values", [])
    if not rows:
        return []

    bairros: List[str] = []
    for idx, row in enumerate(rows):
        nome = (row[0] if row else "").strip()
        if not nome:
            continue
        if idx == 0 and _normalize(nome) in {"nome do bairro", "bairro"}:
            # ignora o cabeçalho
            continue
        bairros.append(nome)
    return bairros


def append_ticket_and_get_number(service, spreadsheet_id: str, sheet_title: str, row_values: List[str]) -> int:
    """
    Faz append da linha (com Senha vazia) e retorna o número da senha atribuído com base no índice da linha.
    Estratégia: append → extrair 'updatedRange' → calcular row_idx → Senha = row_idx - 1 → update célula A{row_idx}
    """
    # Garante a aba e cabeçalhos
    ensure_area_sheet(service, spreadsheet_id, sheet_title)

    # Append mantém a coluna Senha vazia (índice 0) para ser atualizada após o append
    body = {"values": [row_values]}
    append_result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_title}!A1",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body=body,
    ).execute()

    updated_range = (append_result or {}).get("updates", {}).get("updatedRange", "")
    # extrai o número da última linha gravada (mesma técnica usada em utilidades similares)
    m = re.search(r"!.*?(\d+):", updated_range) or re.search(r"!.*?(\d+)$", updated_range)
    if not m:
        raise RuntimeError(f"Não foi possível detectar a linha inserida: {updated_range}")
    row_idx_int = int(m.group(1))

    # Cabeçalho está na linha 1 → senha = row_idx - 1
    senha_num = max(1, row_idx_int - 1)

    # Atualiza a célula A{row_idx} com o número da senha
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_title}!A{row_idx_int}",
        valueInputOption="RAW",
        body={"values": [[str(senha_num)]]},
    ).execute()

    return senha_num


def now_str(tz_name: str = DEFAULT_TIMEZONE) -> str:
    try:
        tz = ZoneInfo(tz_name)
    except Exception:
        tz = None
    dt = datetime.now(tz=tz)
    return dt.strftime("%d/%m/%Y %H:%M:%S")


def _init_ticket_pdf() -> FPDF:
    pdf = FPDF(unit="mm", format=(80, 150))  # ticket com espaço extra para logo/rodapé
    pdf.set_auto_page_break(auto=True, margin=1)
    pdf.set_left_margin(6)
    pdf.set_right_margin(6)
    pdf.set_top_margin(1)
    return pdf


@lru_cache(maxsize=1)
def _resolve_logo_path() -> Optional[Path]:
    candidates: List[Path] = []
    if PDF_LOGO_PATH:
        raw = Path(PDF_LOGO_PATH).expanduser()
        candidates.append(raw)
        if not raw.is_absolute():
            candidates.append(Path(__file__).resolve().parent / raw)
    else:
        candidates.append(DEFAULT_LOGO_PATH)

    for candidate in candidates:
        try:
            if candidate.is_file():
                return candidate
        except OSError:
            continue
    return None


@lru_cache(maxsize=1)
def _load_logo_bytes() -> Optional[bytes]:
    path = _resolve_logo_path()
    if not path:
        return None
    try:
        return path.read_bytes()
    except OSError:
        return None


def _render_ticket_page(pdf: FPDF, data: Dict[str, str]) -> None:
    area = str(data.get("area", "Área")).strip()
    senha = str(data.get("senha", "0")).strip()
    nome = format_name_upper(data.get("nome", ""))
    tel = format_phone_number(data.get("telefone", ""))
    bairro = str(data.get("bairro", "")).strip()
    ts = str(data.get("ts_registro", "")).strip()

    # QR (conteúdo: "AREA|SENHA|NOME")
    qr_payload = f"{area}|{senha}|{nome}"
    qr_img = qrcode.make(qr_payload)
    buf_qr = io.BytesIO()
    qr_img.save(buf_qr, format="PNG")
    buf_qr.seek(0)

    # Code128 com a senha
    buf_bar = io.BytesIO()
    Code128(senha, writer=ImageWriter()).write(buf_bar, options={
        "module_width": 0.3,
        "module_height": 12,
        "font_size": 8,
    })
    buf_bar.seek(0)

    pdf.add_page()

    logo_bytes = _load_logo_bytes()
    if logo_bytes:
        buf_logo = io.BytesIO(logo_bytes)
        buf_logo.name = "logo.png"
        logo_width = 36
        logo_x = (80 - logo_width) / 2
        y_before = pdf.get_y()
        pdf.image(buf_logo, x=logo_x, y=y_before, w=logo_width)
        pdf.set_y(y_before + logo_width + 2)

    # Cabeçalho
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 5, "Distribuidor de Senhas", ln=True, align="C")
    pdf.set_font("Helvetica", "", 12)
    pdf.cell(0, 5, area, ln=True, align="C")
    pdf.ln(1)

    # Senha grande
    pdf.set_font("Helvetica", "B", 40)
    pdf.cell(0, 16, f"{senha}", ln=True, align="C")
    pdf.ln(1)

    # Barra + QR
    x = pdf.get_x()
    y = pdf.get_y()
    pdf.image(buf_bar, x=x + 10, y=y, w=50)
    pdf.ln(18)
    pdf.image(buf_qr, x=(80 - 30) / 2, y=pdf.get_y() + 2, w=30)
    pdf.ln(30)

    # Dados do participante
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 5, f"Nome: {nome}", ln=True)
    pdf.cell(0, 5, f"Telefone: {tel}", ln=True)
    pdf.cell(0, 5, f"Bairro: {bairro}", ln=True)
    pdf.cell(0, 5, f"Registro: {ts}", ln=True)
    pdf.ln(3)

    # Rodapé
    pdf.set_font("Helvetica", "I", 8)
    pdf.multi_cell(0, 4.5, "Guarde este ticket até o atendimento.", align="C")


def _pdf_bytes(pdf: FPDF) -> bytes:
    raw = pdf.output(dest="S")
    return bytes(raw) if isinstance(raw, (bytes, bytearray)) else str(raw).encode("latin-1")


def generate_ticket_pdf(data: Dict[str, str]) -> bytes:
    """Gera um PDF de ticket único."""

    pdf = _init_ticket_pdf()
    _render_ticket_page(pdf, data)
    return _pdf_bytes(pdf)


def generate_tickets_pdf(tickets: List[Dict[str, str]]) -> bytes:
    """Gera um PDF com uma página por ticket."""

    pdf = _init_ticket_pdf()
    for ticket in tickets:
        _render_ticket_page(pdf, ticket)
    return _pdf_bytes(pdf)


def submit_tickets(
    areas: List[str],
    nome: str,
    telefone: str,
    bairro: str,
    rede_social: str = "",
    email: str = "",
) -> Tuple[List[Dict[str, Any]], Optional[bytes], List[Dict[str, Any]]]:
    """Submete uma ou mais senhas para diferentes áreas."""

    if not areas:
        raise ValueError("Selecione ao menos uma área ativa.")

    service = _sheets_service()
    spreadsheet_id = _get_spreadsheet_id()

    # Consulta áreas ativas e mapeamento area->sheet
    areas_info = read_active_areas(service, spreadsheet_id)
    map_area_sheet = {a["area"]: a["sheet"] for a in areas_info}
    map_area_limit = {a["area"]: a.get("max_senhas") for a in areas_info}

    nome_fmt = format_name_upper(nome)
    if not nome_fmt:
        raise ValueError("Nome é obrigatório.")

    telefone_fmt = format_phone_number(telefone)
    bairro_fmt = (bairro or "").strip()
    rede_social_fmt = (rede_social or "").strip()
    email_fmt = (email or "").strip()

    resultados: List[Dict[str, Any]] = []
    tickets_payload: List[Dict[str, str]] = []
    excedidas: List[Dict[str, Any]] = []

    for area in areas:
        sheet_title = map_area_sheet.get(area) or area
        ts = now_str()
        row = [
            "",
            nome_fmt,
            telefone_fmt,
            rede_social_fmt,
            email_fmt,
            bairro_fmt,
            ts,
            "",
        ]
        senha_num = append_ticket_and_get_number(service, spreadsheet_id, sheet_title, row)

        registro = {
            "area": area,
            "sheet": sheet_title,
            "senha": str(senha_num),
            "nome": nome_fmt,
            "telefone": telefone_fmt,
            "bairro": bairro_fmt,
            "rede_social": rede_social_fmt,
            "email": email_fmt,
            "ts_registro": ts,
        }
        resultados.append(registro)
        tickets_payload.append(registro)

        limite = map_area_limit.get(area)
        if limite is not None and senha_num > limite:
            excedidas.append({
                "area": area,
                "limite": limite,
                "senha": senha_num,
            })

    pdf_bytes: Optional[bytes] = None
    if not excedidas:
        pdf_bytes = generate_tickets_pdf(tickets_payload)
    return resultados, pdf_bytes, excedidas


def submit_ticket(
    area: str,
    nome: str,
    telefone: str,
    bairro: str,
    rede_social: str = "",
    email: str = "",
) -> Tuple[int, bytes]:
    """Compatibilidade: submete apenas uma área."""

    resultados, pdf_bytes, excedidas = submit_tickets(
        [area], nome, telefone, bairro, rede_social=rede_social, email=email
    )
    if not resultados:
        raise ValueError("Falha ao gerar a senha.")
    if excedidas:
        info = excedidas[0]
        raise ValueError(
            f"A área {info['area']} excedeu o limite de {info['limite']} senhas (atual: {info['senha']})."
        )
    return int(resultados[0]["senha"]), pdf_bytes
