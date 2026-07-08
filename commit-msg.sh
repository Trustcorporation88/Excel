cd /c/Users/meifl/agente-excel || exit 1
git add .
git commit --trailer "Co-authored-by: Cursor <cursoragent@cursor.com>" -m "$(cat <<'EOF'
Initialize Agente Excel app with DeepSeek integration and guided question flows.

Set up the React/Vite project, add Excel analysis modes, and support suggest/auto/free handling for user d?vidas.
EOF
)"
git status --short
