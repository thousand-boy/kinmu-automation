"""入力テンプレートの構造と勤務ルール。"""

SCHEDULE_SHEET_NAME = "Sheet1"
MASTER_SHEET_NAME = "職員マスタ"
DATE_ROW = 2
WEEKDAY_ROW = 3
DAY_START_COLUMN = 3  # C
DAY_END_COLUMN = 33  # AG

MASTER_HEADERS = (
    "職員ID", "表示名", "病棟", "専従", "専任", "作業療法士", "在籍",
)
WORK_VALUES = {None, "休", "出", "半/", "/半"}
REASON_VALUES = {
    None, "希望", "委員会", "年休", "年/", "/年", "研修", "特休", "出張",
}
FIXED_WORK_VALUES = WORK_VALUES - {None}
FIXED_REASON_VALUES = REASON_VALUES - {None}
WEEKDAYS = {"月", "火", "水", "木", "金", "土", "日"}
MON_TO_SAT = WEEKDAYS - {"日"}

# 月・火・土の病棟別目安: (最低人数, 最大人数)。Noneは上限なし。
WARD_STAFF_LIMITS = {1: (2, 3), 2: (4, 4), 3: (3, None)}

WEIGHT_LEADER = 200
WEIGHT_OT_MINIMUM = 200
WEIGHT_TOO_MANY_OFF = 120
WEIGHT_THIRD_OFF = 40
WEIGHT_WARD_STAFF = 100
WEIGHT_EXISTING_SHIFT = 20
WEIGHT_OT_SECOND = 10
WEIGHT_ISOLATED_WORK = 2
WEIGHT_FRAGMENTED_WORK = 1
