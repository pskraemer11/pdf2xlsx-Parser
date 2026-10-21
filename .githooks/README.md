# Git-Hooks: Datenschutz-Schleuse

Diese Hooks setzen die verbindliche Datenschutzregel des Repos um
(Regel-Quelle: `lib.sh`). Gesperrte E-Mail-Domains, persönliche Daten
(unter `data/`, `privat/` bzw. mit geschützten Endungen, plus lokale
`COMPACTIONS.md`) und die Bank-Kennung erreichen nie einen Commit oder
Push — und Diagnosen drucken gefundene Werte aus Datenschutzgründen
nie aus.

## Installation (lokal)

    git config core.hooksPath .githooks

## Prüfregeln

| Hook         | Prüfung                                                            |
|--------------|--------------------------------------------------------------------|
| `pre-commit` | Identität, neue Textzeilen, Datenpfade im Staging                  |
| `commit-msg` | Text der Commit-Nachricht                                          |
| `pre-push`   | alle neuen Commits: Identität, Nachricht, Dateiinhalte, Datenpfade |

## Notfall

Ein dokumentierter Bypass ist möglich:

    RELEASE_ALLOW=1 git push …

Nutze ihn bewusst und nur in echten Ausnahmen — niemals als Routine.
