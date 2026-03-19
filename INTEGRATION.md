# Odyssey Bridge Integration

## Что уже реализовано

- Owlbear extension теперь хранит расширенные свойства токена:
  - `identity.playerId`
  - `identity.characterId`
  - `lastRoll`
  - `history`
  - `sync.lastEventId`
- GM-клиент Owlbear:
  - подтягивает roll events из внешнего bridge API
  - применяет их к токену
  - перерисовывает оверлеи на игровом столе
  - пушит актуальные состояния tracked-токенов обратно в bridge
- Python bridge (`bridge_server.py`) хранит:
  - состояние токенов в SQLite
  - очередь событий бросков

## Поток данных

1. Discord-бот рассчитывает бросок.
2. Бот отправляет roll event в `POST /api/rolls`.
3. Owlbear extension у GM опрашивает `GET /api/events`.
4. Extension применяет `effects` к токену:
   - `minor_delta`
   - `serious_delta`
   - изменения `body`
   - identity-поля
5. Изменённый токен сразу отражается на столе через attachment overlays.
6. Extension отправляет актуальный snapshot токена в `POST /api/tokens/upsert`.

## Запуск bridge

```powershell
python D:\Documents\Odyssey\Owlbear\owlbear-hp-main\bridge_server.py
```

По умолчанию bridge слушает:

- `http://127.0.0.1:8765`

Файл базы:

- `D:\Documents\Odyssey\Owlbear\owlbear-hp-main\bridge_state.db`

## Формат roll event

```json
{
  "room_id": "owlbear-room-id",
  "event": {
    "token_id": "scene-token-id",
    "actor_name": "Odyssey Bot",
    "summary": "Torso hit for 2 damage",
    "payload": {
      "total": 87,
      "outcome": "success",
      "target_part": "Torso",
      "effects": {
        "minor_delta": 1,
        "body": {
          "Torso": {
            "current_delta": -2
          }
        }
      }
    }
  }
}
```

## Поддерживаемые `effects`

- `minor`
- `minor_delta`
- `serious`
- `serious_delta`
- `identity.player_id`
- `identity.character_id`
- `body.<PART>.current`
- `body.<PART>.current_delta`
- `body.<PART>.max`
- `body.<PART>.max_delta`
- `body.<PART>.armor`
- `body.<PART>.armor_delta`

Где `<PART>`:

- `Head`
- `Torso`
- `L.Arm`
- `R.Arm`
- `L.Leg`
- `R.Leg`

## Интеграция с Python-ботом

В workspace добавлен helper:

- `D:\Documents\Odyssey\odyssey_bridge_client.py`

Пример:

```python
from odyssey_bridge_client import send_roll_event

send_roll_event(
    room_id="owlbear-room-id",
    token_id="scene-token-id",
    actor_name="Odyssey Bot",
    summary="Headshot for 1 damage",
    total=96,
    outcome="critical-success",
    target_part="Head",
    effects={
        "body": {
            "Head": {"current_delta": -1}
        }
    },
)
```

## Что осталось сделать следующим шагом

- Подключить `Odyssey_roll.py` к `send_roll_event(...)`
- Сопоставить Discord-персонажа с `token_id` Owlbear
- При желании добавить отдельную таблицу персонажей/игроков и авторизацию GM

## Установка в Owlbear

1. Размести папку `D:\Documents\Odyssey\Owlbear\owlbear-hp-main` на статическом HTTPS-хостинге.
2. Убедись, что по URL доступны:
   - `manifest.json`
   - `index.html`
   - `background.html`
   - `assets/main.js`
   - `assets/background.js`
3. В Owlbear Rodeo открой `Extensions`.
4. Нажми `Install from manifest URL`.
5. Вставь прямую HTTPS-ссылку на `manifest.json`.

## Быстрая проверка v1

1. Зайди в комнату как GM.
2. Добавь character token на сцену.
3. Открой расширение и нажми `Track Character`.
4. В блоке `Odyssey ownership` укажи:
   - `Owner Player ID`
   - `Owner Name`
5. В блоке `Odyssey stats` задай минимум:
   - один боевой skill
   - `Parry`
   - `Default weapon damage`
6. В блоке `Odyssey actions` выбери:
   - skill
   - часть тела цели
   - модификаторы
7. Нажми `Attack` или `Roll Dice`.

Ожидаемое поведение:

- `GM` может бросать за любой tracked token.
- Игрок может бросать только за токен, где `owner.playerId` совпадает с его Owlbear player id.
- После `Attack` обновляется `lastRoll`.
- При попадании уменьшается `current HP` выбранной части тела.
- Overlay на столе обновляется автоматически.
