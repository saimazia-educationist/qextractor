"""
Accounts - email/password login, credit balances, and manual payment
verification (Easypaisa / bank transfer), backed by a small SQLite database.

Kept as a separate module (rather than folded into app.py) so the
auth/credits/payments logic is easy to find and test on its own.
"""
import re
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

from werkzeug.security import generate_password_hash, check_password_hash

# How many free credits a brand-new account starts with, and how many
# credits each paid action (an extraction run, or a test build) costs.
# Change these to match your pricing.
FREE_SIGNUP_CREDITS = 3
EXTRACT_COST = 1
TEST_BUILD_COST = 1

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

_db_path: Path | None = None


class AuthError(Exception):
    """Raised for any user-facing account/payment problem (bad password,
    duplicate email, unknown payment id, etc). Callers can str() it and
    show it directly to the user."""


def init_db(db_path: Path):
    """Creates the users/payments tables if they don't exist yet. Call once
    at app startup, before any other function in this module is used."""
    global _db_path
    _db_path = db_path
    with _get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                credits INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS payments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL REFERENCES users(id),
                method TEXT NOT NULL,
                amount TEXT,
                transaction_id TEXT,
                note TEXT,
                proof_filename TEXT,
                status TEXT NOT NULL DEFAULT 'pending',
                credits_granted INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                reviewed_at TEXT
            )
        """)


@contextmanager
def _get_conn():
    conn = sqlite3.connect(str(_db_path))
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


# --- Users -------------------------------------------------------------

def register(email: str, password: str) -> dict:
    email = (email or "").strip().lower()
    if not EMAIL_RE.match(email):
        raise AuthError("Please enter a valid email address.")
    if not password or len(password) < 6:
        raise AuthError("Password must be at least 6 characters.")

    with _get_conn() as conn:
        if conn.execute("SELECT 1 FROM users WHERE email = ?", (email,)).fetchone():
            raise AuthError("An account with that email already exists. Try logging in instead.")
        cur = conn.execute(
            "INSERT INTO users (email, password_hash, credits, created_at) VALUES (?, ?, ?, ?)",
            (email, generate_password_hash(password), FREE_SIGNUP_CREDITS, _now()),
        )
        user_id = cur.lastrowid

    return get_user(user_id)


def login(email: str, password: str) -> dict:
    email = (email or "").strip().lower()
    with _get_conn() as conn:
        row = conn.execute("SELECT * FROM users WHERE email = ?", (email,)).fetchone()
    if not row or not check_password_hash(row["password_hash"], password or ""):
        raise AuthError("Incorrect email or password.")
    return dict(row)


def get_user(user_id: int) -> dict | None:
    with _get_conn() as conn:
        row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    return dict(row) if row else None


def public_user(user: dict) -> dict:
    """Strips the password hash before sending a user record to the browser."""
    return {"id": user["id"], "email": user["email"], "credits": user["credits"]}


def try_spend_credits(user_id: int, amount: int) -> bool:
    """Atomically deducts `amount` credits if (and only if) the user has
    enough. Returns True if the spend succeeded, False if they were short."""
    with _get_conn() as conn:
        cur = conn.execute(
            "UPDATE users SET credits = credits - ? WHERE id = ? AND credits >= ?",
            (amount, user_id, amount),
        )
        return cur.rowcount > 0


def refund_credits(user_id: int, amount: int):
    """Used to give credits back if a paid action (extraction/build) fails
    after the credit was already spent."""
    with _get_conn() as conn:
        conn.execute("UPDATE users SET credits = credits + ? WHERE id = ?", (amount, user_id))


# --- Payments ------------------------------------------------------------

VALID_METHODS = {"easypaisa", "bank_transfer"}


def create_payment(user_id: int, method: str, amount: str, transaction_id: str,
                    note: str, proof_filename: str | None) -> int:
    if method not in VALID_METHODS:
        raise AuthError("Unknown payment method.")
    if not transaction_id or not transaction_id.strip():
        raise AuthError("Please enter the transaction ID / reference number from your payment.")

    with _get_conn() as conn:
        cur = conn.execute(
            """INSERT INTO payments
               (user_id, method, amount, transaction_id, note, proof_filename, status, created_at)
               VALUES (?, ?, ?, ?, ?, ?, 'pending', ?)""",
            (user_id, method, (amount or "").strip(), transaction_id.strip(),
             (note or "").strip(), proof_filename, _now()),
        )
        return cur.lastrowid


def list_payments_for_user(user_id: int) -> list[dict]:
    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM payments WHERE user_id = ? ORDER BY id DESC", (user_id,)
        ).fetchall()
    return [dict(r) for r in rows]


def list_payments(status: str | None = None) -> list[dict]:
    """For the admin panel. Joins in the user's email for display."""
    with _get_conn() as conn:
        if status:
            rows = conn.execute(
                """SELECT payments.*, users.email AS user_email FROM payments
                   JOIN users ON users.id = payments.user_id
                   WHERE payments.status = ? ORDER BY payments.id DESC""",
                (status,),
            ).fetchall()
        else:
            rows = conn.execute(
                """SELECT payments.*, users.email AS user_email FROM payments
                   JOIN users ON users.id = payments.user_id
                   ORDER BY payments.id DESC"""
            ).fetchall()
    return [dict(r) for r in rows]


def get_payment(payment_id: int) -> dict | None:
    with _get_conn() as conn:
        row = conn.execute("SELECT * FROM payments WHERE id = ?", (payment_id,)).fetchone()
    return dict(row) if row else None


def review_payment(payment_id: int, approve: bool, credits_to_grant: int = 0) -> dict:
    """Approves or rejects a pending payment. On approval, credits the
    submitting user's account. Raises AuthError if the payment doesn't
    exist or was already reviewed (so an admin can't double-click their
    way into granting credits twice)."""
    with _get_conn() as conn:
        payment = conn.execute("SELECT * FROM payments WHERE id = ?", (payment_id,)).fetchone()
        if not payment:
            raise AuthError("Payment not found.")
        if payment["status"] != "pending":
            raise AuthError("This payment has already been reviewed.")

        if approve:
            if credits_to_grant <= 0:
                raise AuthError("Enter how many credits to grant before approving.")
            conn.execute(
                "UPDATE payments SET status='approved', credits_granted=?, reviewed_at=? WHERE id=?",
                (credits_to_grant, _now(), payment_id),
            )
            conn.execute(
                "UPDATE users SET credits = credits + ? WHERE id = ?",
                (credits_to_grant, payment["user_id"]),
            )
        else:
            conn.execute(
                "UPDATE payments SET status='rejected', credits_granted=0, reviewed_at=? WHERE id=?",
                (_now(), payment_id),
            )

    return get_payment(payment_id)
