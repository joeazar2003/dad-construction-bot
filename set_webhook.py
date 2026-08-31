"""
Run this once per bot, after that service is deployed and live on Render, so
Telegram knows where to send its messages.

Usage:
    python set_webhook.py <bot-token> <https://site1-expenses-bot.onrender.com>

Run it 3 times total, once per site, with that site's own token and URL.
"""
import sys

import requests

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python set_webhook.py <bot-token> <https://your-service.onrender.com>")
        sys.exit(1)

    token, base_url = sys.argv[1], sys.argv[2]
    url = f"{base_url.rstrip('/')}/webhook"
    resp = requests.post(f"https://api.telegram.org/bot{token}/setWebhook", json={"url": url})
    print(resp.json())
