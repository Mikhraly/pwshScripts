$bom = Import-Csv bom.csv

$bom[0].Designator.Replace(" ", "") -split(",")


# $word = New-Object -ComObject Word.Application
# $word.Visible = $True
# $doc = $word.Documents.Open("C:\Users\gasratov\Desktop\bom.docx")
# $table = $doc.Content.Tables.Item(1)
# $table.Select()
# $table.Cell(4,7).Range.Text = "ЯКНИ.01"
