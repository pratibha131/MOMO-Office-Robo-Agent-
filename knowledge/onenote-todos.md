# OneNote to-do board (desktop OneNote COM — tested recipes)

The user's visible to-do board lives in OneNote: notebook **"Pratibha @ Philips"** →
section **"Momo"** → page **"Momo To-Dos"** (already created). data/todos.json stays the
machine-readable source of truth; OneNote is the human mirror. KEEP THEM IN SYNC:
- ADD a todo → append a checkbox line to the OneNote page AND an entry in todos.json.
- COMPLETE a todo → tick + strike through the OneNote line (NEVER delete it — struck
  lines are the record) AND set done:true in todos.json.

## Boilerplate
```powershell
$on = New-Object -ComObject OneNote.Application
$ns = "http://schemas.microsoft.com/office/onenote/2013/onenote"
$hier = ""; $on.GetHierarchy("", 4, [ref]$hier); [xml]$h = $hier   # 4 = hsPages
$page = $h.SelectNodes("//*[local-name()='Page']") | Where-Object { $_.name -eq 'Momo To-Dos' } | Select-Object -First 1
$pageId = $page.ID
# read:  $c = ""; $on.GetPageContent($pageId, [ref]$c, 0); [xml]$doc = $c
# write: $on.UpdatePageContent($doc.OuterXml)
```
If the page is ever missing: find section 'Momo' under the first notebook, else
`$on.OpenHierarchy("Momo.one", $nb.ID, [ref]$sectionId, 3)` to create the section,
`$on.CreateNewPage($sectionId, [ref]$pageId, 0)`, then seed with a one:Title
"Momo To-Dos" and `<one:TagDef index="0" type="0" symbol="3" fontColor="automatic"
highlightColor="none" name="To Do"/>` (symbol 3 = checkbox).

## Add a task (checkbox line)
```powershell
$nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable); $nsm.AddNamespace("one", $ns)
$children = $doc.SelectSingleNode("//one:Outline/one:OEChildren", $nsm)
$oe = $doc.CreateElement("one", "OE", $ns)
$tag = $doc.CreateElement("one", "Tag", $ns)
$tag.SetAttribute("index", "0"); $tag.SetAttribute("completed", "false")
$t = $doc.CreateElement("one", "T", $ns)
$null = $t.AppendChild($doc.CreateCDataSection("Fix the PCP execution sheet Excel file  (due 2 Sep, for Rahul)"))
$null = $oe.AppendChild($tag); $null = $oe.AppendChild($t); $null = $children.AppendChild($oe)
$on.UpdatePageContent($doc.OuterXml)
```
Line format: task text, then two spaces + brackets with due/assigner when known.

## Complete a task (tick the box + strike the text — KEEP the line)
```powershell
foreach ($oe in $doc.SelectNodes("//one:Outline//one:OE", $nsm)) {
  $tnode = $oe.SelectSingleNode("one:T", $nsm)
  if ($tnode -and $tnode.InnerText -like "*PCP execution*") {     # fuzzy-match the task
    $tagNode = $oe.SelectSingleNode("one:Tag", $nsm)
    if ($tagNode) { $tagNode.SetAttribute("completed", "true") }
    $plain = $tnode.InnerText -replace '<[^>]+>', ''
    $tnode.InnerXml = ""
    $null = $tnode.AppendChild($doc.CreateCDataSection(
      "<span style='text-decoration:line-through'>$plain</span>"))
  }
}
$on.UpdatePageContent($doc.OuterXml)
```
Match the task the user described fuzzily (they will paraphrase). If nothing matches or
2+ lines match, ask the user which one — never guess a strike.

## Notes
- one:T CDATA accepts inline HTML: the line-through span renders as strikethrough.
- Escape any `<`, `&` in task text before putting it in CDATA-HTML (or strip them).
- After changes, optionally `$on.NavigateTo($pageId, "", $false)` to show the page —
  do this when the user asked to SEE their to-dos, not on every background sync.
- If COM fails ("cannot create object"), desktop OneNote may not be running/installed —
  say so and fall back to todos.json only.
