import os
import re
import json
import uuid
import base64
import logging
from datetime import datetime
from zoneinfo import ZoneInfo
from collections import defaultdict

import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

load_dotenv()

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)

BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
TIMEZONE = os.getenv("TIMEZONE", "Asia/Makassar")
ALLOWED_USER_IDS_RAW = os.getenv("ALLOWED_USER_IDS", "").strip()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

TRANSACTION_HEADERS = [
    "ID",
    "Timestamp",
    "Tanggal",
    "Bulan",
    "User ID",
    "Username",
    "Tipe",
    "Kategori",
    "Deskripsi",
    "Jumlah",
]

DAILY_HEADERS = ["Tanggal", "Pemasukan", "Pengeluaran", "Saldo"]
MONTHLY_HEADERS = ["Bulan", "Pemasukan", "Pengeluaran", "Saldo"]
CATEGORY_HEADERS = ["Bulan", "Tipe", "Kategori", "Total"]

if not BOT_TOKEN:
    raise RuntimeError("TELEGRAM_BOT_TOKEN belum diisi.")

if not SPREADSHEET_ID:
    raise RuntimeError("SPREADSHEET_ID belum diisi.")


def get_allowed_user_ids() -> set[int]:
    if not ALLOWED_USER_IDS_RAW:
        return set()

    result = set()
    for item in ALLOWED_USER_IDS_RAW.split(","):
        item = item.strip()
        if item:
            result.add(int(item))
    return result


ALLOWED_USER_IDS = get_allowed_user_ids()


def is_allowed(user_id: int) -> bool:
    if not ALLOWED_USER_IDS:
        return True
    return user_id in ALLOWED_USER_IDS


def rupiah(value: int | float) -> str:
    return "Rp{:,.0f}".format(value).replace(",", ".")


def now_local() -> datetime:
    return datetime.now(ZoneInfo(TIMEZONE))


def get_credentials():
    """
    Mendukung 3 cara kredensial:
    1. Local file service_account.json
    2. Environment GOOGLE_SERVICE_ACCOUNT_JSON berisi JSON mentah
    3. Environment GOOGLE_SERVICE_ACCOUNT_JSON berisi base64 dari JSON
    """
    if SERVICE_ACCOUNT_JSON:
        try:
            data = json.loads(SERVICE_ACCOUNT_JSON)
        except json.JSONDecodeError:
            decoded = base64.b64decode(SERVICE_ACCOUNT_JSON).decode("utf-8")
            data = json.loads(decoded)

        return Credentials.from_service_account_info(data, scopes=SCOPES)

    return Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)


def get_spreadsheet():
    credentials = get_credentials()
    client = gspread.authorize(credentials)
    return client.open_by_key(SPREADSHEET_ID)


def get_or_create_worksheet(spreadsheet, title: str, headers: list[str], rows: int = 1000, cols: int = 20):
    try:
        worksheet = spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)
        worksheet.append_row(headers)
        return worksheet

    existing = worksheet.row_values(1)
    if existing != headers:
        if not existing:
            worksheet.append_row(headers)
        else:
            worksheet.update("A1", [headers])

    return worksheet


def setup_sheets():
    spreadsheet = get_spreadsheet()
    transaksi = get_or_create_worksheet(spreadsheet, "Transaksi", TRANSACTION_HEADERS)
    rekap_harian = get_or_create_worksheet(spreadsheet, "Rekap_Harian", DAILY_HEADERS)
    rekap_bulanan = get_or_create_worksheet(spreadsheet, "Rekap_Bulanan", MONTHLY_HEADERS)
    rekap_kategori = get_or_create_worksheet(spreadsheet, "Rekap_Kategori", CATEGORY_HEADERS)
    return spreadsheet, transaksi, rekap_harian, rekap_bulanan, rekap_kategori


def parse_money(text: str) -> int | None:
    cleaned = text.lower().strip()
    cleaned = cleaned.replace("rp", "")
    cleaned = cleaned.replace("idr", "")
    cleaned = cleaned.replace(" ", "")

    multiplier = 1

    if cleaned.endswith("k"):
        multiplier = 1000
        cleaned = cleaned[:-1]
    elif cleaned.endswith("rb"):
        multiplier = 1000
        cleaned = cleaned[:-2]
    elif cleaned.endswith("jt"):
        multiplier = 1000000
        cleaned = cleaned[:-2]

    cleaned = cleaned.replace(".", "").replace(",", "")

    if not cleaned.isdigit():
        return None

    return int(cleaned) * multiplier


def parse_transaction_text(text: str, default_type: str = "pengeluaran"):
    """
    Format:
    /catat 25000 makan siang #makan
    25k kopi #minuman
    /masuk 500000 gaji freelance #income
    """
    text = text.strip()

    text = re.sub(r"^/(catat|masuk|pengeluaran|pemasukan)(@\w+)?", "", text, flags=re.IGNORECASE).strip()

    if not text:
        return None

    parts = text.split()
    amount = parse_money(parts[0])

    if amount is None or amount <= 0:
        return None

    rest = " ".join(parts[1:]).strip()

    category = "lainnya"
    category_match = re.search(r"#([\w\-]+)", rest)

    if category_match:
        category = category_match.group(1).lower()
        description = re.sub(r"#([\w\-]+)", "", rest).strip()
    else:
        description = rest

    if not description:
        description = "-"

    return {
        "tipe": default_type,
        "jumlah": amount,
        "kategori": category,
        "deskripsi": description,
    }


def read_transactions():
    _, transaksi_ws, _, _, _ = setup_sheets()
    rows = transaksi_ws.get_all_records()
    return rows, transaksi_ws


def append_transaction(user, parsed: dict):
    _, transaksi_ws, _, _, _ = setup_sheets()

    current = now_local()
    tx_id = str(uuid.uuid4())[:8]
    tanggal = current.strftime("%Y-%m-%d")
    bulan = current.strftime("%Y-%m")
    timestamp = current.strftime("%Y-%m-%d %H:%M:%S")

    username = user.username or user.full_name or "-"

    row = [
        tx_id,
        timestamp,
        tanggal,
        bulan,
        str(user.id),
        username,
        parsed["tipe"],
        parsed["kategori"],
        parsed["deskripsi"],
        parsed["jumlah"],
    ]

    transaksi_ws.append_row(row, value_input_option="USER_ENTERED")
    rebuild_recaps()

    return {
        "id": tx_id,
        "timestamp": timestamp,
        "tanggal": tanggal,
        "bulan": bulan,
        "username": username,
        **parsed,
    }


def rebuild_recaps():
    spreadsheet, transaksi_ws, harian_ws, bulanan_ws, kategori_ws = setup_sheets()
    rows = transaksi_ws.get_all_records()

    daily = defaultdict(lambda: {"pemasukan": 0, "pengeluaran": 0})
    monthly = defaultdict(lambda: {"pemasukan": 0, "pengeluaran": 0})
    category = defaultdict(int)

    for row in rows:
        try:
            tanggal = str(row.get("Tanggal", "")).strip()
            bulan = str(row.get("Bulan", "")).strip()
            tipe = str(row.get("Tipe", "")).strip().lower()
            kategori = str(row.get("Kategori", "lainnya")).strip().lower()
            jumlah = int(float(row.get("Jumlah", 0)))
        except Exception:
            continue

        if tipe not in ["pemasukan", "pengeluaran"]:
            continue

        if tanggal:
            daily[tanggal][tipe] += jumlah

        if bulan:
            monthly[bulan][tipe] += jumlah
            category[(bulan, tipe, kategori)] += jumlah

    daily_values = [DAILY_HEADERS]
    for tanggal in sorted(daily.keys()):
        pemasukan = daily[tanggal]["pemasukan"]
        pengeluaran = daily[tanggal]["pengeluaran"]
        saldo = pemasukan - pengeluaran
        daily_values.append([tanggal, pemasukan, pengeluaran, saldo])

    monthly_values = [MONTHLY_HEADERS]
    for bulan in sorted(monthly.keys()):
        pemasukan = monthly[bulan]["pemasukan"]
        pengeluaran = monthly[bulan]["pengeluaran"]
        saldo = pemasukan - pengeluaran
        monthly_values.append([bulan, pemasukan, pengeluaran, saldo])

    category_values = [CATEGORY_HEADERS]
    for (bulan, tipe, kategori_name), total in sorted(category.items()):
        category_values.append([bulan, tipe, kategori_name, total])

    harian_ws.clear()
    harian_ws.update("A1", daily_values)

    bulanan_ws.clear()
    bulanan_ws.update("A1", monthly_values)

    kategori_ws.clear()
    kategori_ws.update("A1", category_values)


def summarize_period(period_type: str, target: str):
    rows, _ = read_transactions()

    pemasukan = 0
    pengeluaran = 0
    by_category = defaultdict(int)
    count = 0

    period_column = "Tanggal" if period_type == "hari" else "Bulan"

    for row in rows:
        if str(row.get(period_column, "")).strip() != target:
            continue

        tipe = str(row.get("Tipe", "")).strip().lower()
        kategori = str(row.get("Kategori", "lainnya")).strip().lower()

        try:
            jumlah = int(float(row.get("Jumlah", 0)))
        except Exception:
            continue

        count += 1

        if tipe == "pemasukan":
            pemasukan += jumlah
        elif tipe == "pengeluaran":
            pengeluaran += jumlah
            by_category[kategori] += jumlah

    saldo = pemasukan - pengeluaran

    return {
        "target": target,
        "count": count,
        "pemasukan": pemasukan,
        "pengeluaran": pengeluaran,
        "saldo": saldo,
        "by_category": dict(sorted(by_category.items(), key=lambda item: item[1], reverse=True)),
    }


def summarize_categories_current_month():
    current_month = now_local().strftime("%Y-%m")
    rows, _ = read_transactions()

    by_category = defaultdict(int)

    for row in rows:
        if str(row.get("Bulan", "")).strip() != current_month:
            continue

        tipe = str(row.get("Tipe", "")).strip().lower()
        if tipe != "pengeluaran":
            continue

        kategori = str(row.get("Kategori", "lainnya")).strip().lower()

        try:
            jumlah = int(float(row.get("Jumlah", 0)))
        except Exception:
            continue

        by_category[kategori] += jumlah

    return current_month, dict(sorted(by_category.items(), key=lambda item: item[1], reverse=True))


def get_total_balance():
    rows, _ = read_transactions()

    pemasukan = 0
    pengeluaran = 0

    for row in rows:
        tipe = str(row.get("Tipe", "")).strip().lower()

        try:
            jumlah = int(float(row.get("Jumlah", 0)))
        except Exception:
            continue

        if tipe == "pemasukan":
            pemasukan += jumlah
        elif tipe == "pengeluaran":
            pengeluaran += jumlah

    return pemasukan, pengeluaran, pemasukan - pengeluaran


def format_summary(title: str, summary: dict) -> str:
    lines = [
        f"📊 {title}",
        "",
        f"Jumlah transaksi: {summary['count']}",
        f"Total pemasukan: {rupiah(summary['pemasukan'])}",
        f"Total pengeluaran: {rupiah(summary['pengeluaran'])}",
        f"Saldo: {rupiah(summary['saldo'])}",
    ]

    if summary["by_category"]:
        lines.append("")
        lines.append("Pengeluaran per kategori:")
        for kategori, total in summary["by_category"].items():
            lines.append(f"- #{kategori}: {rupiah(total)}")

    return "\n".join(lines)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    if not is_allowed(user.id):
        await update.message.reply_text("Maaf, kamu tidak diizinkan memakai bot ini.")
        return

    text = f"""
Halo, {user.first_name}! 👋

Bot siap mencatat keuangan harianmu.

User ID kamu:
{user.id}

Contoh catat pengeluaran:
/catat 25000 makan siang #makan

Atau langsung:
25000 kopi #minuman

Catat pemasukan:
/masuk 500000 gaji freelance #income

Rekap:
/hari
/bulan
/kategori
/saldo

Hapus transaksi terakhir kamu:
/hapus
""".strip()

    await update.message.reply_text(text)


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = """
Panduan format:

1. Catat pengeluaran
/catat 25000 makan siang #makan
/catat 15k kopi #minuman
/catat 1jt bayar kontrakan #rumah

2. Catat pemasukan
/masuk 500000 gaji freelance #income

3. Input cepat
25000 nasi goreng #makan

4. Rekap
/hari - rekap hari ini
/bulan - rekap bulan ini
/kategori - kategori pengeluaran bulan ini
/saldo - total pemasukan, pengeluaran, dan saldo
/hapus - hapus transaksi terakhir milikmu

Kategori ditulis dengan tanda pagar, misalnya:
#makan
#transport
#belanja
#tagihan
#hiburan
#kesehatan
""".strip()

    await update.message.reply_text(text)


async def catat(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    if not is_allowed(user.id):
        await update.message.reply_text("Maaf, kamu tidak diizinkan memakai bot ini.")
        return

    parsed = parse_transaction_text(update.message.text, default_type="pengeluaran")

    if not parsed:
        await update.message.reply_text(
            "Format belum benar.\n\nContoh:\n/catat 25000 makan siang #makan"
        )
        return

    tx = append_transaction(user, parsed)

    await update.message.reply_text(
        "\n".join([
            "✅ Pengeluaran dicatat.",
            "",
            f"ID: {tx['id']}",
            f"Tanggal: {tx['timestamp']}",
            f"Kategori: #{tx['kategori']}",
            f"Deskripsi: {tx['deskripsi']}",
            f"Jumlah: {rupiah(tx['jumlah'])}",
        ])
    )


async def masuk(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    if not is_allowed(user.id):
        await update.message.reply_text("Maaf, kamu tidak diizinkan memakai bot ini.")
        return

    parsed = parse_transaction_text(update.message.text, default_type="pemasukan")

    if not parsed:
        await update.message.reply_text(
            "Format belum benar.\n\nContoh:\n/masuk 500000 gaji freelance #income"
        )
        return

    tx = append_transaction(user, parsed)

    await update.message.reply_text(
        "\n".join([
            "✅ Pemasukan dicatat.",
            "",
            f"ID: {tx['id']}",
            f"Tanggal: {tx['timestamp']}",
            f"Kategori: #{tx['kategori']}",
            f"Deskripsi: {tx['deskripsi']}",
            f"Jumlah: {rupiah(tx['jumlah'])}",
        ])
    )


async def quick_add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    if not is_allowed(user.id):
        await update.message.reply_text("Maaf, kamu tidak diizinkan memakai bot ini.")
        return

    parsed = parse_transaction_text(update.message.text, default_type="pengeluaran")

    if not parsed:
        await update.message.reply_text(
            "Saya belum paham formatnya.\n\nContoh:\n25000 kopi #minuman"
        )
        return

    tx = append_transaction(user, parsed)

    await update.message.reply_text(
        f"✅ Dicatat: {rupiah(tx['jumlah'])} untuk {tx['deskripsi']} #{tx['kategori']}"
    )


async def hari(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    if not is_allowed(user.id):
        await update.message.reply_text("Maaf, kamu tidak diizinkan memakai bot ini.")
        return

    today = now_local().strftime("%Y-%m-%d")
    summary = summarize_period("hari", today)

    await update.message.reply_text(format_summary(f"Rekap hari ini ({today})", summary))


async def bulan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    if not is_allowed(user.id):
        await update.message.reply_text("Maaf, kamu tidak diizinkan memakai bot ini.")
        return

    current_month = now_local().strftime("%Y-%m")
    summary = summarize_period("bulan", current_month)

    await update.message.reply_text(format_summary(f"Rekap bulan ini ({current_month})", summary))


async def kategori(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    if not is_allowed(user.id):
        await update.message.reply_text("Maaf, kamu tidak diizinkan memakai bot ini.")
        return

    current_month, data = summarize_categories_current_month()

    if not data:
        await update.message.reply_text(f"Belum ada pengeluaran pada bulan {current_month}.")
        return

    lines = [f"🏷️ Kategori pengeluaran bulan {current_month}", ""]

    for name, total in data.items():
        lines.append(f"- #{name}: {rupiah(total)}")

    await update.message.reply_text("\n".join(lines))


async def saldo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    if not is_allowed(user.id):
        await update.message.reply_text("Maaf, kamu tidak diizinkan memakai bot ini.")
        return

    pemasukan, pengeluaran, balance = get_total_balance()

    await update.message.reply_text(
        "\n".join([
            "💰 Total saldo",
            "",
            f"Total pemasukan: {rupiah(pemasukan)}",
            f"Total pengeluaran: {rupiah(pengeluaran)}",
            f"Saldo: {rupiah(balance)}",
        ])
    )


async def hapus(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    if not is_allowed(user.id):
        await update.message.reply_text("Maaf, kamu tidak diizinkan memakai bot ini.")
        return

    rows, transaksi_ws = read_transactions()

    if not rows:
        await update.message.reply_text("Belum ada transaksi untuk dihapus.")
        return

    user_id = str(user.id)

    # Cari transaksi terakhir milik user.
    # Header ada di row 1, data mulai row 2.
    for index in range(len(rows) - 1, -1, -1):
        row = rows[index]
        if str(row.get("User ID", "")).strip() == user_id:
            sheet_row_number = index + 2
            desc = row.get("Deskripsi", "-")
            amount = row.get("Jumlah", 0)
            transaksi_ws.delete_rows(sheet_row_number)
            rebuild_recaps()

            await update.message.reply_text(
                f"🗑️ Transaksi terakhir dihapus:\n{desc} - {rupiah(int(float(amount)))}"
            )
            return

    await update.message.reply_text("Tidak ada transaksi milikmu yang bisa dihapus.")


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    logging.exception("Terjadi error:", exc_info=context.error)

    if isinstance(update, Update) and update.effective_message:
        await update.effective_message.reply_text(
            "Maaf, terjadi error. Coba lagi sebentar lagi."
        )


def main():
    setup_sheets()

    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("catat", catat))
    app.add_handler(CommandHandler("pengeluaran", catat))
    app.add_handler(CommandHandler("masuk", masuk))
    app.add_handler(CommandHandler("pemasukan", masuk))
    app.add_handler(CommandHandler("hari", hari))
    app.add_handler(CommandHandler("bulan", bulan))
    app.add_handler(CommandHandler("kategori", kategori))
    app.add_handler(CommandHandler("saldo", saldo))
    app.add_handler(CommandHandler("hapus", hapus))

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, quick_add))

    app.add_error_handler(error_handler)

    logging.info("Bot berjalan...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()