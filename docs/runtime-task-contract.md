# Runtime Task Contract（唯一規格，凍結版）

> 狀態：**Frozen**（v1）  
> 適用範圍：前端（主控）與後端（執行器）  
> 目的：作為前後端**唯一依據**，禁止各自解讀。

---

## 0. 名詞與總則

### 0.1 固定術語（不可替換）
- **Origin**：執行前快照（只讀）。
- **Runtime Working**：每回合由 Origin 複製出的臨時計算結果（不可回寫 Origin）。
- **Execution Progress**：後端回報的執行進度狀態。
- **rslot**：來自 `randat<slot>` 的單列 slot。
- **bslot**：每個 `buff_group` 代表 slot（取群首 idx）。
- **pass**：略過（不觸發按放）。
- **walk**：經過不按（游標經過但不觸發按放）。
- **none**：正常按放。

### 0.2 常數
- 前後端皆必須定義：`RUNTIME_CONTRACT_VERSION = "v1"`。

### 0.3 時間基準
- 所有 `*_ms` 欄位為毫秒。
- 延遲與 timeout 計算使用**單調時鐘**（monotonic clock）。

---

## A. 角色與責任切分（強制）

### A.1 前端（主控）
前端負責：
1. 建立 **Origin Snapshot**。
2. 每回合建立 **Runtime Working**。
3. 在 Runtime Working 上執行：抽籤、換位、顏色標示。
4. 產生最小任務（timeline）。
5. 顯示 `current_idx` 與 runtime JSON（必須可對照後端進度）。

### A.2 後端（純執行器）
後端只負責：
1. 接收任務。
2. 立即執行。
3. 回報進度與狀態。

後端**禁止**：
- 抽籤。
- 重算 `at_ms`。
- 重排 `idx`。

### A.3 buff skip 語意（後端必遵守）
- `skip_mode = "pass"`：略過該列。
- `skip_mode = "walk"`：經過不按。
- `skip_mode = "none"`：正常按放。

---

## B. 三層狀態模型（強制）

### B.1 Origin（只讀）
- 建立時機：執行前。
- 用途：
  1. 恢復到執行前狀態。
  2. 每回合重算起點。
- 禁止被 Runtime 或執行回報污染。

### B.2 Runtime Working（臨時）
- 每回合由 Origin 全量複製建立。
- 僅供本回合抽籤、搬移、著色、排程。
- 不得寫入 undo/redo 歷史。

### B.3 Execution Progress
- 來源：後端事件流。
- 最小欄位：`server_task_id`, `round`, `current_idx`, `state`, `event_time_ms`。
- 前端顯示必須與後端一致，不得自行推測覆蓋。

---

## C. slot 與抽籤規則（強制）

### C.1 slot 名詞與來源
1. `rslot`：每個 `randat<slot>` 產生一個候選 slot（單列）。
2. `bslot`：每個 `buff_group` 只提供一個候選 slot，定義為群首 idx。

### C.2 抽籤池
- 固定為：`pool = rslot + bslot`。

### C.3 抽籤順序
- `buff_group` 依 group id 升序（1,2,3,...) 各抽一次。

### C.4 公平性
- 每次抽中的 slot 必須從 pool 移除。
- 已移除 slot 不可重複使用。

### C.5 顏色語意
- 抽到自己的 `bslot`：維持淺黃。
- 抽到他人的 slot：顯示淺藍。

### C.6 單一 block 規則
- `buff_group` 即使有多列，抽籤時視為**單一 block（one block / one bslot）**。

---

## D. buff_group 搬移與間距規則（強制）

### D.1 搬移定位
- 以群首 idx 作為定位基準。
- 整群同步移動（不得拆列）。

### D.2 A/B gap 語意
- **A gap**：群前間距。
- **B gap**：群後間距。
- 搬移後必須保留前後鄰接節奏，不得破壞 A/B gap 語意。

### D.3 固定範例（整群搬移）
範例：`idx15~idx19`（同一個 buff_group）抽到 `slot=idx37`。
- 搬移前：
  - group head = idx15
  - group body = idx16~idx19
  - 前後節奏依原 timeline 存在 A/B gap
- 搬移操作：
  - 將 **idx15~idx19 作為一個 block** 搬移到以 idx37 為定位點的新位置。
  - 不是只搬單列 idx15。
- 搬移後：
  - 新 group head 對齊 idx37 對應位置。
  - idx16~idx19 保持相對順序與群內距離。
  - 新位置前後的 A/B gap 語意仍成立。

### D.4 寫入限制
- 搬移計算僅發生於 Runtime Working。
- 禁止回寫 Origin。

---

## E. 合法性檢核條件（強制）

### E.1 啟用條件
- 只有 `rslot_count > 0` 才啟用抽籤模式與相關檢核（包含 buff_group 連續性）。

### E.2 關閉條件
- 當 `rslot_count == 0`：
  - 不進抽籤模式。
  - 不套用該組 bslot 連續性強檢核。

### E.3 buff_group 連續性提示
- 若抽籤檢核啟用且發現 buff_group 非連續（中間夾其他列）：
  - 執行前必須提示。
  - 單次執行與重複執行都必須提示。

---

## F. 控制語意（強制）

1. **Repeat**：每回合皆從 Origin 重新計算 Runtime，再送後端。
2. **Stop**：全部停止；可提示是否恢復 Origin。
3. **Pause**：僅暫停，不清空進度；按鈕文字切為 `Continue`。
4. **Resume**：從暫停點繼續；前後端狀態必須一致。

---

## G. 傳輸契約（JSON Schema 風格，強制）

> 本節為欄位最小集合。可增欄位，但不得刪除必填欄位。

### G.1 Start Task Request

```json
{
  "type": "start_task",
  "contract_version": "v1",
  "client_task_id": "cli-20260421-0001",
  "round": 1,
  "delay_ms": 2000,
  "sent_at_ms": 1234567890,
  "origin_version": 7,
  "runtime_version": 23,
  "timeline": [
    {
      "idx": 37,
      "at_ms": 0,
      "action": "tap",
      "btn": "A",
      "skip_mode": "none"
    }
  ],
  "runtime_meta": {
    "rslot": [5, 37, 90],
    "bslot": [15, 42],
    "pool_initial": [5, 37, 90, 15, 42],
    "draw_order": [1, 2],
    "draw_result": [
      {"group_id": 1, "picked_slot": 37, "picked_from": "rslot", "color": "light_blue"},
      {"group_id": 2, "picked_slot": 42, "picked_from": "bslot", "color": "light_yellow"}
    ]
  }
}
```

必含欄位：
- `contract_version`
- `client_task_id`
- `round`
- `delay_ms`
- `sent_at_ms`
- `origin_version`
- `runtime_version`
- `timeline[]`（每列至少：`idx`, `at_ms`, `action`, `btn`, `skip_mode`）
- `runtime_meta`（抽籤輸入與結果）

### G.2 Ack Response

```json
{
  "type": "ack",
  "server_task_id": "srv-9f8b",
  "state": "accepted",
  "ack_at_ms": 1234567902
}
```

必含欄位：
- `server_task_id`
- `state = accepted`
- `ack_at_ms`

### G.3 Progress / State / Error 事件最小集

**Progress**
```json
{
  "type": "progress",
  "server_task_id": "srv-9f8b",
  "round": 1,
  "current_idx": 37,
  "state": "running",
  "event_time_ms": 1234568123
}
```

**State**
```json
{
  "type": "state",
  "server_task_id": "srv-9f8b",
  "state": "paused",
  "event_time_ms": 1234569000
}
```

允許 state 枚舉：
`accepted | running | paused | resumed | stopped | finished | error`

**Error**
```json
{
  "type": "error",
  "server_task_id": "srv-9f8b",
  "phase": "start",
  "code": "START_TIMEOUT",
  "message": "ack received but running not entered in threshold",
  "diag": {
    "expected_state": "running",
    "last_state": "accepted",
    "elapsed_ms": 5000,
    "threshold_ms": 5000
  },
  "event_time_ms": 1234570000
}
```

---

## H. Timeout 診斷模型（強制）

### H.1 `ACK_TIMEOUT`
- 判定條件：送出 `start_task` 後，於門檻內未收到 `ack`。
- 建議門檻：`1500 ms`。
- 前端顯示文案：`任務已送出，但後端尚未回覆受理（ACK Timeout）。`
- 後端 `diag` 最小欄位：`{"phase":"ack","elapsed_ms":x,"threshold_ms":1500}`。

### H.2 `START_TIMEOUT`
- 判定條件：已收到 `ack`，但門檻內未進入 `running`。
- 建議門檻：`5000 ms`。
- 前端顯示文案：`任務已受理，但未在預期時間內開始執行（Start Timeout）。`
- 後端 `diag` 最小欄位：`{"phase":"start","last_state":"accepted","elapsed_ms":x,"threshold_ms":5000}`。

### H.3 `PROGRESS_STALL`
- 判定條件：進入 `running` 後，連續超過門檻時間無 progress 事件。
- 建議門檻：`3000 ms`（可依任務密度調整）。
- 前端顯示文案：`執行中進度長時間未更新（Progress Stall）。`
- 後端 `diag` 最小欄位：`{"phase":"running","last_progress_at_ms":a,"now_ms":b,"stall_ms":b-a,"threshold_ms":3000}`。

---

## I. 準時送出（延遲 x 秒）規範（強制）

1. 前端必須以單調時鐘作為延遲計算基準，避免系統時鐘跳動影響。
2. 到點立即送出 `start_task`。
3. request 必帶 `sent_at_ms`，供端到端延遲比對。
4. 可接受延遲誤差：
   - 建議指標：`p95 |actual_send_time - scheduled_send_time| <= 30 ms`。
5. 量測方法：
   - 每次送出記錄 `scheduled_send_time_ms` 與 `actual_send_time_ms`（皆單調時鐘）。
   - 每 100 次統計一次 p95，超標需告警。

---

## J. 契約版本鎖（強制）

### J.1 固定版本
- 前後端常數：`RUNTIME_CONTRACT_VERSION = "v1"`。

### J.2 版本不一致處理
- start 時雙方都帶版本。
- 若版本不一致，後端必須拒收並回傳可讀錯誤（含 expected/actual）。

範例：
```json
{
  "type": "error",
  "phase": "validate_contract",
  "code": "CONTRACT_VERSION_MISMATCH",
  "message": "runtime contract version mismatch",
  "diag": {
    "expected": "v1",
    "actual": "v2"
  }
}
```

### J.3 版本升級策略（v1 -> v2）
- **不兼容變更**：
  - 採 hard reject（拒收）策略。
  - 需升級雙端至同版本後再通訊。
- **兼容新增欄位**：
  - 可在 `v1` 下增加 optional 欄位。
  - 不得改動既有必填語意。
- 發布流程建議：
  1. 先後端支援 v1+v2（灰度）。
  2. 前端切流到 v2。
  3. 觀察穩定後移除 v1（另行公告）。

---

## K. 完整 JSON 範例（強制）

### K.1 範例 A：一般 round（rslot+bslot 抽籤、群組搬移、顏色、最終 timeline）

```json
{
  "start_request": {
    "type": "start_task",
    "contract_version": "v1",
    "client_task_id": "cli-20260421-0101",
    "round": 3,
    "delay_ms": 1000,
    "sent_at_ms": 991000,
    "origin_version": 12,
    "runtime_version": 77,
    "timeline": [
      {"idx": 10, "at_ms": 0, "action": "tap", "btn": "A", "skip_mode": "none"},
      {"idx": 11, "at_ms": 120, "action": "tap", "btn": "B", "skip_mode": "none"},
      {"idx": 37, "at_ms": 240, "action": "tap", "btn": "A", "skip_mode": "none"},
      {"idx": 38, "at_ms": 360, "action": "tap", "btn": "B", "skip_mode": "walk"},
      {"idx": 39, "at_ms": 480, "action": "tap", "btn": "A", "skip_mode": "pass"}
    ],
    "runtime_meta": {
      "rslot": [5, 37, 90],
      "bslot": [15, 42],
      "pool_initial": [5, 37, 90, 15, 42],
      "draw_order": [1, 2],
      "draw_result": [
        {
          "group_id": 1,
          "group_range_before": [15, 19],
          "picked_slot": 37,
          "picked_from": "rslot",
          "color": "light_blue",
          "move_mode": "whole_group"
        },
        {
          "group_id": 2,
          "group_range_before": [42, 44],
          "picked_slot": 42,
          "picked_from": "bslot",
          "color": "light_yellow",
          "move_mode": "whole_group"
        }
      ],
      "pool_after_each_draw": [
        [5, 90, 15, 42],
        [5, 90, 15]
      ],
      "gap_rule": {
        "A_gap_preserved": true,
        "B_gap_preserved": true
      }
    }
  },
  "ack": {
    "type": "ack",
    "server_task_id": "srv-0101",
    "state": "accepted",
    "ack_at_ms": 991010
  },
  "events": [
    {
      "type": "state",
      "server_task_id": "srv-0101",
      "state": "running",
      "event_time_ms": 991040
    },
    {
      "type": "progress",
      "server_task_id": "srv-0101",
      "round": 3,
      "current_idx": 10,
      "state": "running",
      "event_time_ms": 991045
    },
    {
      "type": "progress",
      "server_task_id": "srv-0101",
      "round": 3,
      "current_idx": 37,
      "state": "running",
      "event_time_ms": 991280
    },
    {
      "type": "state",
      "server_task_id": "srv-0101",
      "state": "finished",
      "event_time_ms": 991800
    }
  ]
}
```

### K.2 範例 B：Stop 後恢復 Origin，再啟動新 round

```json
{
  "round_1": {
    "start_request": {
      "type": "start_task",
      "contract_version": "v1",
      "client_task_id": "cli-20260421-0201",
      "round": 1,
      "delay_ms": 0,
      "sent_at_ms": 1200000,
      "origin_version": 20,
      "runtime_version": 101,
      "timeline": [
        {"idx": 1, "at_ms": 0, "action": "tap", "btn": "A", "skip_mode": "none"}
      ],
      "runtime_meta": {
        "rslot": [8],
        "bslot": [15],
        "pool_initial": [8, 15],
        "draw_order": [1],
        "draw_result": [
          {"group_id": 1, "picked_slot": 8, "picked_from": "rslot", "color": "light_blue"}
        ]
      }
    },
    "stop": {
      "type": "state",
      "server_task_id": "srv-0201",
      "state": "stopped",
      "event_time_ms": 1200500
    }
  },
  "recover_origin": {
    "action": "restore_origin",
    "origin_version": 20,
    "runtime_version_discarded": 101,
    "result": "ui_restored_to_origin"
  },
  "round_2": {
    "start_request": {
      "type": "start_task",
      "contract_version": "v1",
      "client_task_id": "cli-20260421-0202",
      "round": 2,
      "delay_ms": 500,
      "sent_at_ms": 1202000,
      "origin_version": 20,
      "runtime_version": 102,
      "timeline": [
        {"idx": 2, "at_ms": 0, "action": "tap", "btn": "B", "skip_mode": "none"}
      ],
      "runtime_meta": {
        "rslot": [9],
        "bslot": [15],
        "pool_initial": [9, 15],
        "draw_order": [1],
        "draw_result": [
          {"group_id": 1, "picked_slot": 15, "picked_from": "bslot", "color": "light_yellow"}
        ]
      }
    }
  }
}
```

重點：
- Stop 後恢復使用同一 `origin_version=20`。
- 新回合 `runtime_version` 需遞增（101 -> 102）。

---

## L. 驗收標準（強制）

1. 任一工程師僅閱讀本文件即可實作前後端，無需口頭補充。
2. 文件術語唯一且不衝突（Origin/Runtime/rslot/bslot/pass/walk）。
3. 可用範例 A/B 逐欄位對照前端送出與後端回報。
4. 可用 timeout 章節直接定位卡點（ack、起跑、執行中）。

---

## 附錄：實作一致性清單（建議）

- [ ] 前端是否每回合由 Origin 深拷貝 Runtime Working。
- [ ] 後端是否完全不改 idx / at_ms / 排序。
- [ ] `pool` 是否確實移除已抽 slot。
- [ ] `rslot_count==0` 是否關閉抽籤與連續性強檢核。
- [ ] Pause/Resume 是否保留同一 `server_task_id` 上下文。
- [ ] timeout diag 是否可直接定位問題相位。

