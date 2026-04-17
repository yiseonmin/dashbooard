# SQLite DB 연결 및 함수
import sqlite3
from datetime import datetime

DB_PATH = "data/calendar.db"
# SQLite DB 초기화
# DB 없는 경우 자동 생성함

# ── DB 초기화 (테이블 없으면 자동 생성) ──
def init_db():
conn = sqlite3.connect(DB_PATH)
	conn.execute("""
		CREATE TABLE IF NOT EXISTS schedules (
			id		INTEGER PRIMARY KEY AUTOINCREMENT,
			title	TEXT NOT NULL,
			start 	TEXT NOT NULL,
			end 	TEXT NOT NULL,
			category	TEXT,
			memo 	TEXT,
			color	TEXT,
			created_at TEXT
		)
	""")
	conn.commit()
	conn.close()

# ── 카테고리별 색상 ──────────────────────
def get_color(category):
	colors = {
		"QA 테스트": "#FF4B4B",
		"팀 미팅": "#4B9FFF",
	}
	return colors.get(category, "#808080")

# ── 일정 등록 ────────────────────────────
def add_schedule(title, start, end, category, memo=""):
	conn = sqlite3.connect(DB_PATH)
	conn.execute("""
		INSERT INTO schedules (title, start, end, category, memo, color, created_at)
		VALUES (?, ?, ?, ?, ?, ?, ?)
	""", (
		title,
		start,
		end,
		category,
		memo,
		get_color(category),
		datetime.now().strftime("%Y-%m-%d %H:%M:%S")
	))
	conn.commit()
	conn.close()


# ── 일정 전체 조회 → streamlit-calendar 형식 ──
def get_all_schedules():
	conn = sqlite3.connect(DB_PATH)
	rows = conn.execute("""
		SELECT id, title, start, end, category, memo, color
		FROM schedules
		ORDER BY start
	""").fetchall()
	conn.close()

	# ✅ SQLite → JSON 형태로 변환
	events = []
	for row in rows:
		events.append({
			"id":  row[0],
			"title": row[1],
			"start": row[2],
			"end":  row[3],
			"color": row[6],
			"extendedProps": {
				"category": row[4],
				"memo": 	row[5]
			}
		})
	return events

# ── 일정 삭제 ────────────────────────────
def delete_schedule(schedule_id):
	conn = sqlite3.connect(DB_PATH)
	conn.execute("DELETE FROM schedules WHERE id = ?", (schedule_id,))
	conn.commit()
	conn.close()
