from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler

# Eski botning tokenini qo'ying (hali mavjud bo'lsa)
OLD_BOT_TOKEN = "8964353995:AAEfLBAygqtrygBXEcFj6lzArVYNwXXJmQs"

async def start(update: Update, context):
    # Faqat bitta tugma - yangi botga havola
    keyboard = [
        [InlineKeyboardButton("🚀 Yangi botga o'tish", url="https://t.me/QuizForgeUzBot")]
    ]
    await update.message.reply_text(
        "⚠️ Ushbu bot faoliyati tugatilgan.\n"
        "Barcha imkoniyatlar yangi botda mavjud. Quyidagi tugma orqali o'ting:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

def main():
    app = Application.builder().token(OLD_BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    print("Eski bot redirect uchun ishga tushdi...")
    app.run_polling()

if __name__ == "__main__":
    main()
