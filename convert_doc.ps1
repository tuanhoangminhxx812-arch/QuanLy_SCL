
 = New-Object -ComObject Word.Application
.Visible = False
 = (Resolve-Path 'HS_mau\Quyết định phê duyệt quyết toán CT SCL-VTAD2607001.doc').Path
 = [System.IO.Path]::ChangeExtension(, '.docx')
 = .Documents.Open()
.SaveAs([ref], [ref]16)
.Close()
.Quit()
Write-Host 'SUCCESS'
