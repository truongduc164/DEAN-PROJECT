"""
TelegramBot – background Telegram bot for remote administration.

Runs in a daemon thread, polls for commands from an authorized Telegram
chat and executes admin actions:

    /changepw <user> <new_password>  – change app password
    /status                          – report hostname & app version
    /ping                            – bot health check
"""
from __future__ import annotations

import logging
import platform
import threading
import time
from typing import Optional

import requests

logger = logging.getLogger("DeanTran.telegram_bot")

# ╔══════════════════════════════════════════════════════════════════╗
# ║  ĐIỀN BOT TOKEN VÀ CHAT ID CỦA BẠN VÀO ĐÂY                   ║
# ╚══════════════════════════════════════════════════════════════════╝
BOT_TOKEN = "8784102439:AAFiZh97O5_2kFpPAogEpI6SD2SkMPbPD-0"                  # ← Paste bot token từ @BotFather
ALLOWED_CHAT_IDS = [6480561673]           # ← Danh sách chat_id, ví dụ: [123456789]


class TelegramBot:
    """Lightweight Telegram bot using raw HTTP long-polling."""

    def __init__(
        self,
    ) -> None:
        self._token: str = BOT_TOKEN
        self._allowed_ids: list[int] = list(ALLOWED_CHAT_IDS)
        self._offset: int = 0
        self._running = False
        self._thread: Optional[threading.Thread] = None
        self._processed_updates: set[int] = set()

    @property
    def is_configured(self) -> bool:
        return bool(self._token) and bool(self._allowed_ids)

    # ── Telegram API helpers ──────────────────────────────────────────

    def _api(self, method: str, **params) -> dict:
        url = f"https://api.telegram.org/bot{self._token}/{method}"
        try:
            resp = requests.post(url, json=params, timeout=60)
            return resp.json()
        except Exception as exc:
            logger.debug("Telegram API error: %s", exc)
            return {}

    def _send(self, chat_id: int, text: str) -> None:
        self._api("sendMessage", chat_id=chat_id, text=text, parse_mode="Markdown")

    def _get_updates(self) -> list[dict]:
        result = self._api(
            "getUpdates",
            offset=self._offset,
            timeout=30,
            allowed_updates=["message"],
        )
        return result.get("result", [])

    # ── Command handlers ──────────────────────────────────────────────

    def _handle_message(self, message: dict) -> bool:
        chat_id = message.get("chat", {}).get("id")
        text = (message.get("text") or "").strip()

        if not chat_id or not text:
            return True

        # Auth check
        if chat_id not in self._allowed_ids:
            self._send(chat_id, "⛔ Unauthorized. Your chat ID: `{}`".format(chat_id))
            logger.warning("Unauthorized Telegram access from chat_id=%s", chat_id)
            return True

        # Parse command
        parts = text.split()
        cmd = parts[0].lower().split("@")[0]  # strip @botname

        if cmd == "/ping":
            self._cmd_ping(chat_id)
            return True
        elif cmd == "/status":
            return self._cmd_status(chat_id, parts[1:])
        elif cmd == "/changepw":
            return self._cmd_change_password(chat_id, parts[1:])
        elif cmd == "/changeapppw":
            return self._cmd_change_app_password(chat_id, parts[1:])
        elif cmd == "/help":
            self._cmd_help(chat_id)
            return True
        else:
            self._send(chat_id, "❓ Unknown command. Use /help")
            return True

    def _cmd_ping(self, chat_id: int) -> None:
        hostname = platform.node()
        self._send(chat_id, f"🏓 Pong! Host: `{hostname}`")

    def _cmd_status(self, chat_id: int, args: list[str]) -> bool | str:
        target_host = args[0].lower() if args else None
        hostname = platform.node().lower()
        is_global = (target_host == "all")
        if target_host and not is_global and target_host != hostname:
            return False

        try:
            from app.version import APP_TITLE, VERSION
            ver = VERSION
            title = APP_TITLE
        except Exception:
            ver = "unknown"
            title = "DeanTrans"
        self._send(
            chat_id,
            f"📊 *Status*\n"
            f"• Host: `{platform.node()}`\n"
            f"• App: {title}\n"
            f"• Version: `{ver}`\n"
            f"• Running: ✅",
        )
        return "global" if is_global else True

    def _cmd_change_password(self, chat_id: int, args: list[str]) -> bool | str:
        if len(args) < 3:
            self._send(
                chat_id,
                "⚠️ Usage: `/changepw <hostname|all> <username> <new_password>`\n"
                "Example: `/changepw all admin MyNewPass123`",
            )
            return True

        target_host = args[0].lower()
        hostname = platform.node().lower()
        is_global = (target_host == "all")
        if not is_global and target_host != hostname:
            return False

        username = args[1]
        new_password = args[2]

        try:
            from app.core.auth_manager import AuthManager
            am = AuthManager()
            if am.change_password(username, new_password):
                self._send(
                    chat_id,
                    f"✅ Password changed!\n"
                    f"• Host: `{platform.node()}`\n"
                    f"• User: `{username}`\n"
                    f"• New password: ||{new_password}||",
                )
                logger.info("Password changed for '%s' via Telegram", username)
            else:
                self._send(chat_id, f"❌ User `{username}` not found on `{platform.node()}`.")
        except Exception as exc:
            self._send(chat_id, f"❌ Error on `{platform.node()}`: {exc}")
            logger.error("changepw error: %s", exc)
        return "global" if is_global else True

    def _cmd_change_app_password(self, chat_id: int, args: list[str]) -> bool | str:
        if len(args) < 2:
            self._send(
                chat_id,
                "⚠️ Usage: `/changeapppw <hostname|all> <new_password>`\n"
                "Example: `/changeapppw all MyNewAppPass`",
            )
            return True

        target_host = args[0].lower()
        hostname = platform.node().lower()
        is_global = (target_host == "all")
        if not is_global and target_host != hostname:
            return False

        new_password = args[1]
        try:
            from app.core.app_password import set_app_password, get_app_password
            old_pw = get_app_password()
            set_app_password(new_password)
            self._send(
                chat_id,
                f"✅ App password changed!\n"
                f"• Host: `{platform.node()}`\n"
                f"• Old: ||{old_pw}||\n"
                f"• New: ||{new_password}||\n"
                f"_Có hiệu lực lần mở app tiếp theo trên máy này_",
            )
            logger.info("App password changed via Telegram")
        except Exception as exc:
            self._send(chat_id, f"❌ Error on `{platform.node()}`: {exc}")
            logger.error("changeapppw error: %s", exc)
        return "global" if is_global else True

    def _cmd_help(self, chat_id: int) -> None:
        self._send(
            chat_id,
            "🤖 *DeanTrans Bot Commands*\n\n"
            "/ping – Check bot is alive\n"
            "/status `<hostname|all>` – Host & app info\n"
            "/changepw `<hostname|all>` `<user>` `<pass>` – Đổi mật khẩu admin\n"
            "/changeapppw `<hostname|all>` `<pass>` – Đổi mật khẩu mở app\n"
            "/help – Hiển thị trợ giúp",
        )

    # ── Polling loop ──────────────────────────────────────────────────

    def _poll_loop(self) -> None:
        logger.info("Telegram bot polling started (token=...%s)", self._token[-6:])
        consecutive_errors = 0

        while self._running:
            try:
                updates = self._get_updates()
                consecutive_errors = 0

                for update in updates:
                    update_id = update["update_id"]
                    msg = update.get("message")
                    if msg:
                        if hasattr(self, "_processed_updates") and update_id in self._processed_updates:
                            consumed = False
                        else:
                            consumed = self._handle_message(msg)
                            if consumed == "global":
                                if not hasattr(self, "_processed_updates"):
                                    self._processed_updates = set()
                                self._processed_updates.add(update_id)
                                consumed = False
                                
                        if consumed:
                            self._offset = update_id + 1
                        else:
                            # Not consumed (likely for another host or global). Check expiration.
                            msg_time = msg.get("date", 0)
                            if time.time() - msg_time > 60:
                                self._offset = update_id + 1
                                chat_id = msg.get("chat", {}).get("id")
                                text = msg.get("text", "")
                                if chat_id and text.startswith("/") and update_id not in getattr(self, "_processed_updates", set()):
                                    self._send(chat_id, f"⚠️ Lệnh `{text}` hết hạn do không tìm thấy hostname mục tiêu.")
                    else:
                        self._offset = update_id + 1

            except Exception as exc:
                consecutive_errors += 1
                wait = min(30, 2 ** consecutive_errors)
                logger.debug("Poll error (%d): %s – retry in %ds", consecutive_errors, exc, wait)
                time.sleep(wait)

    # ── Start / Stop ──────────────────────────────────────────────────

    def start(self) -> bool:
        """Start the bot in a daemon thread. Returns False if not configured."""
        if not self.is_configured:
            logger.info("Telegram bot not configured – skipping.")
            return False

        if self._running:
            return True

        self._running = True
        self._thread = threading.Thread(
            target=self._poll_loop, daemon=True, name="TelegramBot"
        )
        self._thread.start()
        
        # Send startup notification
        if self._allowed_ids:
            try:
                from app.version import APP_TITLE
                app_name = APP_TITLE
                msg = (
                    f"📗 🟢 {app_name} đã khởi động\n"
                    f"🖥 Máy: `{platform.node()}`\n"
                    f"🏷 App: `{app_name}`"
                )
                self._send(self._allowed_ids[0], msg)
            except Exception:
                pass
                
        logger.info("Telegram bot started.")
        return True

    def stop(self) -> None:
        """Stop the polling loop."""
        self._running = False
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=5)
        logger.info("Telegram bot stopped.")
