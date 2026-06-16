-- strip-bookmarks.lua
-- Removes heading anchor bookmarks that pandoc injects into DOCX output.
-- These appear as clickable bookmark icons in Word/Google Docs — unwanted in most docs.
-- Usage: pandoc input.md -o output.docx --lua-filter=SKILL_DIR/filters/strip-bookmarks.lua

function Header(el)
  el.identifier = ""
  return el
end

function Span(el)
  if el.identifier ~= "" then
    return el.content
  end
end
