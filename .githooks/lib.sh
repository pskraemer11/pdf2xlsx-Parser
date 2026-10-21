# Gemeinsame Prüfregeln für die Hooks in .githooks/ (Quelle per ". …/lib.sh").
#
# Verbindliche Datenschutzregel dieses Repos:
#  - gesperrte private E-Mail-Domains (online.de, e-mail.de, posteo.de) kommen
#    nie vor in Autor-/Committer-Identitäten, Commit-Nachrichten, Dateiinhalten
#    oder hinzugefügten Textzeilen
#  - persönliche Daten (die Ordner data/ und privat/ sowie pdf/xlsx/ods/csv/txt/zip
#    und lokale Projekt-Memo COMPACTIONS.md) werden nie Bestandteil eines Commits
#  - die Kennung der betroffenen Bank ("diba") kommt in Dateien/Nachrichten nie vor
# Diagnosen geben gefundene Werte aus Datenschutzgründen NIE aus – verweigert
# wird pauschal mit Verweis auf die Regel.

MUSTER_DOMAINS='online\.de|e-mail\.de|posteo\.de'
MUSTER_BANK='diba'
MUSTER_DATEN='(^data/|^privat/|(^|/)COMPACTIONS\.md$|\.(pdf|xlsx|ods|csv|txt|zip)$)'

verweigern() {
    echo "❌ $1" >&2
    echo "   Regelquelle: .githooks/lib.sh — dokumentierter Notfall: RELEASE_ALLOW=1 (bewusst nutzen, nie zur Routine)." >&2
    exit 1
}

hook_bypass_aktiv() {
    [ "${RELEASE_ALLOW:-0}" = "1" ]
}
