import asyncio
import os

from aiogram import Bot, Dispatcher
from aiogram.fsm.storage.memory import MemoryStorage
from dotenv import load_dotenv

from access import AccessManager
from handlers import router
from render import ensure_assets_exist, register_unbounded_fonts, DATA_DIR, ACCESS_FILE


async def main():
    load_dotenv()

    token = os.getenv("BOT_TOKEN", "").strip()
    if not token:
        raise RuntimeError("В .env нет BOT_TOKEN")

    admin_ids_raw = os.getenv("ADMIN_IDS", "").strip()
    admin_ids = []
    if admin_ids_raw:
        for x in admin_ids_raw.split(","):
            x = x.strip()
            if x.isdigit():
                admin_ids.append(int(x))

    ensure_assets_exist()
    fonts = register_unbounded_fonts()

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    access = AccessManager(ACCESS_FILE, admin_ids=admin_ids)

    bot = Bot(token=token)
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)

    # прокидываем зависимости в хэндлеры
    await dp.start_polling(bot, access=access, fonts=fonts)


if __name__ == "__main__":
    asyncio.run(main())
