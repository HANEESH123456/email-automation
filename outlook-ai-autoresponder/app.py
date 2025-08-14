import json
import os
import re
import sys
import time
import html
import logging
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional, Tuple

import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
import msal

try:
	from openai import OpenAI, AzureOpenAI
except Exception:
	OpenAI = None
	AzureOpenAI = None


# -------------------- Logging --------------------
logging.basicConfig(
	level=logging.INFO,
	format="%(asctime)s %(levelname)s %(message)s",
)
logger = logging.getLogger("outlook_ai_autoresponder")


# -------------------- Config --------------------
load_dotenv()

AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID", "").strip()
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "").strip()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY", "").strip()
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT", "").strip()
AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION", "2024-06-01").strip()

LLM_MODEL = os.getenv("LLM_MODEL", "gpt-4o-mini").strip()

POLL_INTERVAL_SECONDS = int(os.getenv("POLL_INTERVAL_SECONDS", "30"))
STATE_FILE = os.getenv("STATE_FILE", "state.json").strip()
TOKEN_CACHE_FILE = os.getenv("TOKEN_CACHE_FILE", "token_cache.json").strip()
ALLOWED_LANGUAGES = [lang.strip() for lang in os.getenv("ALLOWED_LANGUAGES", "").split(",") if lang.strip()]

USER_NAME = os.getenv("USER_NAME", "").strip()
USER_ROLE = os.getenv("USER_ROLE", "").strip()
ORG_NAME = os.getenv("ORG_NAME", "").strip()
REPLY_SIGNATURE = os.getenv("REPLY_SIGNATURE", "").strip()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = [
	"Mail.ReadWrite",
	"Mail.Send",
	"offline_access",
]

CATEGORY_NAME = "AutoReplied"
REPLIED_IDS_MAX = 500


# -------------------- Utilities --------------------

def now_utc_iso() -> str:
	return datetime.now(timezone.utc).isoformat()


def parse_iso(dt_str: str) -> datetime:
	return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))


def strip_html_to_text(html_content: str) -> str:
	try:
		soup = BeautifulSoup(html_content or "", "html.parser")
		text = soup.get_text("\n")
		return re.sub(r"\n{3,}", "\n\n", text).strip()
	except Exception:
		return re.sub(r"<[^>]+>", " ", html_content or "").strip()


def load_state(path: str) -> Dict:
	if not os.path.exists(path):
		return {"last_checked_iso": now_utc_iso(), "replied_ids": []}
	with open(path, "r", encoding="utf-8") as f:
		return json.load(f)


def save_state(path: str, state: Dict) -> None:
	with open(path, "w", encoding="utf-8") as f:
		json.dump(state, f, indent=2)


# -------------------- LLM --------------------
class LLMClient:
	def __init__(self):
		self.provider = None
		if AZURE_OPENAI_API_KEY and AZURE_OPENAI_ENDPOINT:
			if AzureOpenAI is None:
				raise RuntimeError("openai package not available. Please install requirements.")
			self.provider = "azure"
			self.client = AzureOpenAI(
				azure_endpoint=AZURE_OPENAI_ENDPOINT,
				api_key=AZURE_OPENAI_API_KEY,
				api_version=AZURE_OPENAI_API_VERSION,
			)
			logger.info("Using Azure OpenAI provider")
		elif OPENAI_API_KEY:
			if OpenAI is None:
				raise RuntimeError("openai package not available. Please install requirements.")
			self.provider = "openai"
			self.client = OpenAI(api_key=OPENAI_API_KEY)
			logger.info("Using OpenAI provider")
		else:
			raise RuntimeError("No LLM provider configured. Set OPENAI_API_KEY or Azure OpenAI envs.")

	def generate_reply(self, prompt_messages: List[Dict], temperature: float = 0.2, max_tokens: int = 600) -> str:
		resp = self.client.chat.completions.create(
			model=LLM_MODEL,
			messages=prompt_messages,
			temperature=temperature,
			max_tokens=max_tokens,
		)
		return resp.choices[0].message.content.strip()


# -------------------- Graph Client --------------------
class GraphClient:
	def __init__(self, tenant_id: str, client_id: str):
		if not tenant_id or not client_id:
			raise ValueError("AZURE_TENANT_ID and AZURE_CLIENT_ID are required")
		self.authority = f"https://login.microsoftonline.com/{tenant_id}"
		self.client_id = client_id
		self.cache = msal.SerializableTokenCache()
		if os.path.exists(TOKEN_CACHE_FILE):
			try:
				self.cache.deserialize(open(TOKEN_CACHE_FILE, "r", encoding="utf-8").read())
			except Exception:
				logger.warning("Could not read token cache; proceeding with empty cache")
		self.msal_app = msal.PublicClientApplication(
			client_id=self.client_id,
			authority=self.authority,
			token_cache=self.cache,
		)
		self.access_token = None
		self.me = None

	def _persist_cache(self) -> None:
		if self.cache.has_state_changed:
			with open(TOKEN_CACHE_FILE, "w", encoding="utf-8") as f:
				f.write(self.cache.serialize())

	def acquire_token(self) -> str:
		accounts = self.msal_app.get_accounts()
		result = None
		if accounts:
			result = self.msal_app.acquire_token_silent(SCOPES, account=accounts[0])
			self._persist_cache()
		if not result:
			flow = self.msal_app.initiate_device_flow(scopes=SCOPES)
			if "user_code" not in flow:
				raise RuntimeError("Failed to create device flow")
			print(f"\nTo sign in, use a browser to open {flow['verification_uri']} and enter the code: {flow['user_code']}\n")
			result = self.msal_app.acquire_token_by_device_flow(flow)
			self._persist_cache()
		if "access_token" not in result:
			raise RuntimeError(f"Auth failed: {result.get('error_description')}")
		self.access_token = result["access_token"]
		return self.access_token

	def _headers(self) -> Dict[str, str]:
		if not self.access_token:
			self.acquire_token()
		return {"Authorization": f"Bearer {self.access_token}", "Content-Type": "application/json"}

	def get_me(self) -> Dict:
		if self.me:
			return self.me
		resp = requests.get(f"{GRAPH_BASE}/me", headers=self._headers(), timeout=30)
		resp.raise_for_status()
		self.me = resp.json()
		return self.me

	def list_recent_messages(self, since_iso: str, top: int = 50, page_limit: int = 5) -> List[Dict]:
		# Use OData params, fetch multiple pages if necessary
		filter_time = parse_iso(since_iso) - timedelta(minutes=2)
		filter_expr = f"receivedDateTime ge {filter_time.astimezone(timezone.utc).isoformat()}"
		select = "id,subject,from,receivedDateTime,conversationId,categories,isRead"
		params = {
			"$filter": filter_expr,
			"$orderby": "receivedDateTime desc",
			"$top": str(top),
			"$select": select,
		}
		url = f"{GRAPH_BASE}/me/mailFolders/inbox/messages"
		headers = self._headers()

		items: List[Dict] = []
		page_count = 0
		while url and page_count < page_limit:
			resp = requests.get(url, headers=headers, params=params if page_count == 0 else None, timeout=30)
			resp.raise_for_status()
			data = resp.json()
			items.extend(data.get("value", []))
			url = data.get("@odata.nextLink")
			page_count += 1
		return items

	def get_message(self, message_id: str) -> Dict:
		select = "id,subject,from,replyTo,toRecipients,ccRecipients,receivedDateTime,conversationId,body,categories"
		url = f"{GRAPH_BASE}/me/messages/{message_id}"
		params = {"$select": select}
		resp = requests.get(url, headers=self._headers(), params=params, timeout=30)
		resp.raise_for_status()
		return resp.json()

	def list_conversation_tail(self, conversation_id: str, limit: int = 5) -> List[Dict]:
		if not conversation_id:
			return []
		select = "id,from,subject,receivedDateTime,bodyPreview"
		url = f"{GRAPH_BASE}/me/messages"
		params = {
			"$filter": f"conversationId eq '{conversation_id}'",
			"$orderby": "receivedDateTime desc",
			"$top": str(limit),
			"$select": select,
		}
		resp = requests.get(url, headers=self._headers(), params=params, timeout=30)
		resp.raise_for_status()
		return resp.json().get("value", [])

	def ensure_category_exists(self, category_name: str) -> None:
		url = f"{GRAPH_BASE}/me/outlook/masterCategories"
		resp = requests.get(url, headers=self._headers(), timeout=30)
		resp.raise_for_status()
		names = {c.get("displayName") for c in resp.json().get("value", [])}
		if category_name not in names:
			create_body = {"displayName": category_name, "color": "preset0"}
			resp2 = requests.post(url, headers=self._headers(), json=create_body, timeout=30)
			if not (200 <= resp2.status_code < 300):
				logger.warning("Failed to create category %s: %s", category_name, resp2.text)

	def add_category_to_message(self, message_id: str, category_name: str) -> None:
		# Merge with existing categories to avoid overwriting
		get_url = f"{GRAPH_BASE}/me/messages/{message_id}"
		resp_get = requests.get(get_url, headers=self._headers(), params={"$select": "categories"}, timeout=30)
		if not (200 <= resp_get.status_code < 300):
			logger.warning("Failed to read categories for %s: %s", message_id, resp_get.text)
			existing = []
		else:
			existing = (resp_get.json() or {}).get("categories", []) or []
		new_cats = list(sorted({*existing, category_name}))
		patch_url = f"{GRAPH_BASE}/me/messages/{message_id}"
		patch = {"categories": new_cats}
		resp = requests.patch(patch_url, headers=self._headers(), json=patch, timeout=30)
		if not (200 <= resp.status_code < 300):
			logger.warning("Failed to set category on %s: %s", message_id, resp.text)

	def reply_to_message(self, message_id: str, html_body: str, dry_run: bool = False) -> None:
		# Create a draft reply, update the body, then send
		create_url = f"{GRAPH_BASE}/me/messages/{message_id}/createReply"
		headers = self._headers()
		resp = requests.post(create_url, headers=headers, timeout=30)
		resp.raise_for_status()
		draft = resp.json()
		draft_id = draft.get("id")
		if not draft_id:
			raise RuntimeError("Failed to create reply draft")

		if dry_run:
			logger.info("[DRY-RUN] Would reply to message %s with body length %d", message_id, len(html_body))
			return

		patch_url = f"{GRAPH_BASE}/me/messages/{draft_id}"
		patch = {"body": {"contentType": "HTML", "content": html_body}}
		resp2 = requests.patch(patch_url, headers=headers, json=patch, timeout=30)
		resp2.raise_for_status()

		send_url = f"{GRAPH_BASE}/me/messages/{draft_id}/send"
		resp3 = requests.post(send_url, headers=headers, timeout=30)
		resp3.raise_for_status()


# -------------------- Content --------------------
SYSTEM_PROMPT = (
	"You are an executive assistant who writes concise, professional email replies. "
	"Respond in the same language as the sender whenever possible. "
	"Keep to 3-6 sentences unless more is essential. "
	"Be courteous, specific, and actionable, and ask for any missing details succinctly."
)


def build_prompt(
	me: Dict,
	sender_name: str,
	sender_email: str,
	subject: str,
	plain_body: str,
	recent_thread_summaries: List[str],
	personalization: Dict[str, str],
	allowed_languages: List[str],
) -> List[Dict]:
	persona_bits = []
	if personalization.get("user_name"):
		persona_bits.append(f"Your name: {personalization['user_name']}")
	if personalization.get("user_role"):
		persona_bits.append(f"Your role: {personalization['user_role']}")
	if personalization.get("org_name"):
		persona_bits.append(f"Organization: {personalization['org_name']}")

	lang_hint = ""
	if allowed_languages:
		lang_hint = f"Only reply in one of these languages if it matches the sender: {', '.join(allowed_languages)}."

	thread_context = "\n".join(recent_thread_summaries) if recent_thread_summaries else "(no prior context)"

	system_message = {"role": "system", "content": SYSTEM_PROMPT}
	user_message = {
		"role": "user",
		"content": (
			f"Sender: {sender_name} <{sender_email}>\n"
			f"Subject: {subject}\n"
			f"Your mailbox principal: {me.get('displayName')} <{me.get('mail') or me.get('userPrincipalName')}>\n"
			f"{'; '.join(persona_bits)}\n"
			f"{lang_hint}\n"
			"Recent conversation tail (newest first):\n"
			f"{thread_context}\n\n"
			"Email body (plain text):\n"
			f"""{plain_body}\n\n"""
			"Compose a concise, helpful, professional reply tailored to the sender."
		),
	}
	return [system_message, user_message]


def wrap_reply_html(body_html: str, signature_html: str) -> str:
	parts = [
		"<div style=\"font-family: -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; font-size: 14px; line-height: 1.45; color: #111\">",
		f"<p>{html.escape(body_html).replace('\n', '<br>')}</p>",
	]
	if signature_html:
		parts.append(f"<div style=\"margin-top:12px; color:#444\">{signature_html}</div>")
	parts.append("<hr style=\"margin:16px 0; border:none; border-top:1px solid #eee\">")
	parts.append("<p style=\"color:#666\">This is an automated response generated by an assistant.</p>")
	parts.append("</div>")
	return "".join(parts)


# -------------------- Heuristics --------------------
AUTOREPLY_SUBJECT_HINTS = [
	"automatic reply",
	"out of office",
	"autoreply",
	"auto reply",
	"vacation reply",
]


def looks_like_auto_generated(sender_email: str, subject: str) -> bool:
	s = (subject or "").lower()
	em = (sender_email or "").lower()
	if any(h in s for h in AUTOREPLY_SUBJECT_HINTS):
		return True
	if re.search(r"no[-\s]?reply", em) or re.search(r"do[-\s]?not[-\s]?reply", em):
		return True
	return False


# -------------------- Runner --------------------
class AutoResponder:
	def __init__(self, dry_run: bool = False):
		self.graph = GraphClient(AZURE_TENANT_ID, AZURE_CLIENT_ID)
		self.llm = LLMClient()
		self.state = load_state(STATE_FILE)
		self.dry_run = dry_run
		self.me = self.graph.get_me()
		self.me_email = (self.me.get("mail") or self.me.get("userPrincipalName") or "").lower()
		self.graph.ensure_category_exists(CATEGORY_NAME)

	def _already_replied(self, message_id: str) -> bool:
		return message_id in set(self.state.get("replied_ids", []))

	def _mark_replied(self, message_id: str) -> None:
		replied_ids: List[str] = self.state.get("replied_ids", [])
		replied_ids.append(message_id)
		if len(replied_ids) > REPLIED_IDS_MAX:
			self.state["replied_ids"] = replied_ids[-REPLIED_IDS_MAX:]
		else:
			self.state["replied_ids"] = replied_ids
		save_state(STATE_FILE, self.state)

	def _summarize_thread_tail(self, items: List[Dict]) -> List[str]:
		summaries = []
		for it in items:
			frm = (it.get("from", {}) or {}).get("emailAddress", {})
			name = frm.get("name") or frm.get("address") or "(unknown)"
			subject = it.get("subject") or ""
			preview = it.get("bodyPreview") or ""
			received = it.get("receivedDateTime") or ""
			summaries.append(f"[{received}] {name}: {subject} — {preview[:200]}")
		return summaries

	def process_once(self) -> None:
		last_checked = self.state.get("last_checked_iso")
		if not last_checked:
			last_checked = now_utc_iso()
			self.state["last_checked_iso"] = last_checked
			save_state(STATE_FILE, self.state)

		logger.info("Checking for messages since %s", last_checked)
		messages = self.graph.list_recent_messages(since_iso=last_checked, top=25)
		processed_any = False

		for msg_stub in messages:
			msg_id = msg_stub["id"]
			if self._already_replied(msg_id):
				continue

			stub_categories = msg_stub.get("categories") or []
			if CATEGORY_NAME in stub_categories:
				# Already handled in a previous run
				self._mark_replied(msg_id)
				continue

			frm = (msg_stub.get("from", {}) or {}).get("emailAddress", {})
			sender_email = (frm.get("address") or "").lower()
			sender_name = frm.get("name") or sender_email
			subject = msg_stub.get("subject") or ""

			if sender_email == self.me_email:
				continue
			if looks_like_auto_generated(sender_email, subject):
				logger.info("Skipping auto-generated looking message from %s: %s", sender_email, subject)
				self._mark_replied(msg_id)  # mark to avoid reprocessing
				continue

			full = self.graph.get_message(msg_id)
			body_html = (full.get("body", {}) or {}).get("content") or ""
			plain = strip_html_to_text(body_html)
			thread_tail = self.graph.list_conversation_tail(full.get("conversationId"), limit=4)
			thread_summaries = self._summarize_thread_tail(thread_tail)

			prompt = build_prompt(
				me=self.me,
				sender_name=sender_name,
				sender_email=sender_email,
				subject=subject,
				plain_body=plain,
				recent_thread_summaries=thread_summaries,
				personalization={
					"user_name": USER_NAME,
					"user_role": USER_ROLE,
					"org_name": ORG_NAME,
				},
				allowed_languages=ALLOWED_LANGUAGES,
			)

			try:
				ai_text = self.llm.generate_reply(prompt)
				wrapped_html = wrap_reply_html(ai_text, REPLY_SIGNATURE)
				self.graph.reply_to_message(msg_id, wrapped_html, dry_run=self.dry_run)
				self.graph.add_category_to_message(msg_id, CATEGORY_NAME)
				self._mark_replied(msg_id)
				processed_any = True
				logger.info("Replied to %s (%s)", sender_email, subject)
			except Exception as e:
				logger.exception("Failed to reply to %s: %s", sender_email, e)

		# Move watermark forward if we checked anything
		self.state["last_checked_iso"] = now_utc_iso()
		save_state(STATE_FILE, self.state)

		if not processed_any:
			logger.info("No new actionable messages.")

	def run_forever(self) -> None:
		while True:
			try:
				self.process_once()
			except Exception as e:
				logger.exception("Top-level iteration error: %s", e)
			time.sleep(POLL_INTERVAL_SECONDS)


def main():
	import argparse
	parser = argparse.ArgumentParser(description="Outlook AI Autoresponder")
	parser.add_argument("--once", action="store_true", help="Process a single iteration and exit")
	parser.add_argument("--dry-run", action="store_true", help="Do not send emails, just log")
	args = parser.parse_args()

	responder = AutoResponder(dry_run=args.dry_run)
	if args.once:
		responder.process_once()
	else:
		responder.run_forever()


if __name__ == "__main__":
	main()