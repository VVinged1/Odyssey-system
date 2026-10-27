# Concurrent editing (1.9.2)

Scene events refresh every open character panel. Previously only the focused persisted field and a limited set of local combat controls survived a panel redraw. Other unsaved inputs, including the add-body-part form, could reset when another GM saved a different character.

The panel now preserves complete local body/weapon/ammunition/perk drafts during redraws. Pending autosaves are tracked by token and field until the SDK write finishes; an older acknowledgement cannot clear a newer edit. Failed autosaves retain the user's input and show a retry message. Local tracker writes are ordered per token, while different tokens remain independent. Selection changes do not copy one character's pending field values into another character.

Regression coverage uses two separate Playwright browser pages connected through a shared mock scene store, with writes limited to the SDK-requested item IDs. It edits different characters, broadcasts remote scene changes, delays one save acknowledgement, checks unfocused drafts and token selection isolation. The test reproduces a blank add-body-part name with the previous production main.js and passes with the fix. Pure tests cover edit revisions, independent token queues and recovery after a rejected write.

This change does not provide multi-user conflict merging when two people intentionally overwrite the same field on the same token; it addresses local form resets and local save ordering.
