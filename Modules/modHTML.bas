option compare database
option explicit

function emailcontentgen(subject as string, title as string, subtitle as string, primarymessage as string, detail1 as string, detail2 as string, detail3 as string, optional appname as string = "", optional appid as string = "") as string

if appid <> "" then
    primarymessage = "<a href = ""\\data\mdbdata\WorkingDB\build\workingdb_commands\openNotification.vbs"">" & primarymessage & "</a>"
end if

emailcontentgen = subject & "," & title & "," & subtitle & "," & primarymessage & "," & detail1 & "," & detail2 & "," & detail3 & "," & appname & "," & appid

end function

function generatehtml(primarytitle as string, primarydetails as string, buttontext as string, _
        detail1 as string, detail2 as string, detail3 as string, _
        optional link as string = "", _
        optional addlines as boolean = false, _
        optional appname as string = "", _
        optional appid as string = "" _
        ) as string

dim headertitle as string
dim greeting as string, htmltext as string

headertitle = appname
greeting = "Hi"

htmltext = htmlheader(primarytitle, 600, addlines) & _
            htmlbody(headertitle, _
                primarytitle, _
                greeting, _
                primarydetails, _
                buttontext, _
                detail1, detail2, detail3, appname, appid, link) & htmlfooter

generatehtml = htmltext

end function

public function genemail(byval strto as string, byval strbcc as string, byval strsubject as string, body as string) as boolean
genemail = true
    
dim objemail as object

set objemail = createobject("outlook.Application")
set objemail = objemail.createitem(0)

with objemail
    .to = strto
    .bcc = strbcc
    .subject = strsubject
    .htmlbody = body
    .display
end with

set objemail = nothing
    
end function

public function htmlheader(title as string, optional customwidth as long = 600, optional addstuff as boolean = false) as string

dim lines as collection
dim i as long

set lines = new collection

lines.add "<!doctype html>"
lines.add "<html lang=""en"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">"
lines.add "<head>"
lines.add "  <meta charset=""utf-8"">"
lines.add "  <meta name=""viewport"" content=""width=device-width, initial-scale=1"">"
lines.add "  <meta name=""x-apple-disable-message-reformatting"">"
lines.add "  <meta name=""color-scheme"" content=""light dark"">"
lines.add "  <meta name=""supported-color-schemes"" content=""light dark"">"
lines.add "  <title>" & title & "</title>"
lines.add ""
lines.add "  <!--[if mso]>"
lines.add "  <noscript>"
lines.add "    <xml>"
lines.add "      <o:OfficeDocumentSettings>"
lines.add "        <o:PixelsPerInch>96</o:PixelsPerInch>"
lines.add "      </o:OfficeDocumentSettings>"
lines.add "    </xml>"
lines.add "  </noscript>"
lines.add "  <![endif]-->"
lines.add ""
lines.add "  <style>"
lines.add "    body, table, td, a { -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%; }"
lines.add "    table, td { mso-table-lspace:0pt; mso-table-rspace:0pt; }"
lines.add "    img { -ms-interpolation-mode:bicubic; border:0; outline:none; text-decoration:none; display:block; }"
lines.add "    table { border-collapse:collapse !important; }"
lines.add "    body { margin:0 !important; padding:0 !important; width:100% !important; background:#f3f4f6; }"
lines.add ""
lines.add "    @media screen and (max-width:600px) {"
lines.add "      .container { width:100% !important; }"
lines.add "      .card-padding { padding:28px 22px !important; }"
lines.add "      .headline { font-size:22px !important; line-height:30px !important; }"
lines.add "      .mobile-full { width:100% !important; }"
lines.add "      .mobile-center { text-align:center !important; }"
lines.add "    }"
lines.add ""
lines.add "    @media (prefers-color-scheme: dark) {"
lines.add "      .email-bg { background:#111827 !important; }"
lines.add "      .card { background:#1f2937 !important; border-color:#374151 !important; }"
lines.add "      .headline { color:#f9fafb !important; }"
lines.add "      .body-text { color:#d1d5db !important; }"
lines.add "      .muted-text { color:#9ca3af !important; }"
lines.add "      .divider { border-top-color:#374151 !important; }"
lines.add "      .brand-text { color:#60a5fa !important; }"
lines.add "    }"
lines.add "  </style>"
lines.add "</head>"
lines.add ""
lines.add "<body style=""margin:0; padding:0; background:#f3f4f6;"">"
lines.add "  <center role=""article"" aria-roledescription=""email"" lang=""en"" style=""width:100%; background:#f3f4f6;"" class=""email-bg"">"
lines.add ""
lines.add "    <!-- Preheader -->"
lines.add "    <div style=""display:none; max-height:0; overflow:hidden; opacity:0; color:transparent; line-height:1px; font-size:1px;"">"
lines.add "      Your notification details are ready."
lines.add "    </div>"
lines.add ""
lines.add "    <table role=""presentation"" width=""100%"" cellpadding=""0"" cellspacing=""0"" class=""email-bg"" style=""background:#f3f4f6;"">"
lines.add "      <tr>"
lines.add "        <td align=""center"" style=""padding:32px 16px;"">"
lines.add ""
lines.add "          <table role=""presentation"" width=""" & customwidth & """ cellpadding=""0"" cellspacing=""0"" class=""container"" style=""width:" & customwidth & "px; max-width:" & customwidth & "px;"">"
lines.add ""
lines.add "            <!-- Header -->"
lines.add "            <tr>"
lines.add "              <td align=""center"" style=""padding:0 0 20px 0;"">"
lines.add "                <div class=""brand-text"" style=""font-family:Segoe UI, Arial, sans-serif; font-size:22px; line-height:28px; font-weight:700; color:#2563eb;"">"
lines.add "                  WorkingDB"
lines.add "                </div>"
lines.add "              </td>"
lines.add "            </tr>"

if addstuff then
    lines.add "            <tr>"
    lines.add "              <td align=""center"" style=""padding:0 0 20px 0;"">"
    lines.add "                <div class=""brand-text"" style=""font-family:Segoe UI, Arial, sans-serif; font-size:18px; line-height:20px; font-weight:600; color:#770000;"">"
    lines.add "                  Extra Notes: type here..."
    lines.add "                </div>"
    lines.add "              </td>"
    lines.add "            </tr>"
end if

for i = 1 to lines.count
    htmlheader = htmlheader & lines(i) & vbcrlf
next i

end function

public function htmlbody( _
        headertitle as string, _
        primarytitle as string, _
        greeting as string, _
        primarydetails as string, _
        buttontext as string, _
        detail1 as string, _
        detail2 as string, _
        detail3 as string, _
        appname as string, _
        appid as string, _
        link as string _
) as string

dim lines as collection
dim i as long

set lines = new collection

lines.add ""
lines.add "            <!-- Card -->"
lines.add "            <tr>"
lines.add "              <td class=""card card-padding"" style=""background:#ffffff; border:1px solid #e5e7eb; border-radius:16px; padding:40px;"">"
lines.add ""
lines.add "                <table role=""presentation"" width=""100%"" cellpadding=""0"" cellspacing=""0"">"
lines.add ""
lines.add "                  <tr>"
lines.add "                    <td>"
lines.add "                      <span style=""display:inline-block; background:#dbeafe; color:#1d4ed8; font-family:Segoe UI, Arial, sans-serif; font-size:12px; line-height:16px; font-weight:700; letter-spacing:.4px; text-transform:uppercase; padding:6px 12px; border-radius:999px;"">"
lines.add "                        " & headertitle & ""
lines.add "                      </span>"
lines.add "                    </td>"
lines.add "                  </tr>"
lines.add ""
lines.add "                  <tr>"
lines.add "                    <td style=""padding-top:24px;"">"
lines.add "                      <h1 class=""headline"" style=""margin:0; font-family:Segoe UI, Arial, sans-serif; font-size:24px; line-height:32px; font-weight:700; color:#111827;"">"
lines.add "                        " & primarytitle & ""
lines.add "                      </h1>"
lines.add "                    </td>"
lines.add "                  </tr>"
lines.add ""
lines.add "                  <tr>"
lines.add "                    <td style=""padding-top:16px;"">"
lines.add "                      <p class=""body-text"" style=""margin:0; font-family:Segoe UI, Arial, sans-serif; font-size:16px; line-height:24px; color:#4b5563;"">"
lines.add "                        " & greeting & ",<br><br>"
lines.add "                        " & primarydetails & ""
lines.add "                      </p>"
lines.add "                    </td>"
lines.add "                  </tr>"
lines.add ""
lines.add "                  <tr>"
lines.add "                    <td style=""padding-top:32px;"">"
lines.add ""
lines.add "                      <!-- Bulletproof Outlook Button -->"
lines.add "                      <table role=""presentation"" cellpadding=""0"" cellspacing=""0"" class=""mobile-full"">"
lines.add "                        <tr>"
lines.add "                          <td align=""center"" style=""border-radius:10px;"">"


if link <> "" then
    lines.add "                            <a href=""" & link & """ style=""display:inline-block; padding:14px 24px; font-family:Segoe UI, Arial, sans-serif; font-size:15px; line-height:20px; font-weight:700; border-radius:10px;"">"
end if

lines.add "                              " & buttontext & ""

if link <> "" then
    lines.add "                            </a>"
end if


lines.add "                          </td>"
lines.add "                        </tr>"
lines.add "                      </table>"
lines.add ""
lines.add "                    </td>"
lines.add "                  </tr>"
lines.add ""
lines.add "                  <tr>"
lines.add "                    <td style=""padding-top:32px;"">"
lines.add "                      <div class=""divider"" style=""border-top:1px solid #e5e7eb; line-height:1px; font-size:1px;"">&nbsp;</div>"
lines.add "                    </td>"
lines.add "                  </tr>"
lines.add ""
lines.add "                  <tr>"
lines.add "                    <td style=""padding-top:24px;"">"
lines.add "                      <p class=""muted-text"" style=""margin:0; font-family:Segoe UI, Arial, sans-serif; font-size:13px; line-height:20px; color:#6b7280;"">"
lines.add "                        " & detail1 & "<br>"
lines.add "                        " & detail2 & "<br>"
lines.add "                        " & detail3 & ""
lines.add "                      </p>"
lines.add "                    </td>"
lines.add "                  </tr>"
lines.add "                  <tr>"
lines.add "                    <td style=""padding-top:24px;"">"
lines.add "                      <p class=""muted-text"" style=""margin:0; font-family:Segoe UI, Arial, sans-serif; font-size:13px; line-height:20px; color:#6b7280;"">"
lines.add "                        AppName:[" & appname & "], AppId:[" & appid & "]"
lines.add "                      </p>"
lines.add "                    </td>"
lines.add "                  </tr>"
lines.add ""
lines.add "                </table>"
lines.add ""
lines.add "              </td>"
lines.add "            </tr>"
lines.add ""

for i = 1 to lines.count
    htmlbody = htmlbody & lines(i) & vbcrlf
next i

end function

public function htmlfooter() as string

dim lines as collection
dim i as long

set lines = new collection

lines.add ""
lines.add "            <!-- Footer -->"
lines.add "            <tr>"
lines.add "              <td align=""center"" style=""padding:24px 20px 0 20px;"">"
lines.add "                <p class=""muted-text"" style=""margin:0; font-family:Segoe UI, Arial, sans-serif; font-size:13px; line-height:20px; color:#6b7280;"">"
lines.add "                  © " & year(date) & " VBA Decoded<br>"
lines.add "                  This is an automated email."
lines.add "                </p>"
lines.add "              </td>"
lines.add "            </tr>"
lines.add ""
lines.add "          </table>"
lines.add ""
lines.add "        </td>"
lines.add "      </tr>"
lines.add "    </table>"
lines.add ""
lines.add "  </center>"
lines.add "</body>"
lines.add "</html>"

for i = 1 to lines.count
    htmlfooter = htmlfooter & lines(i) & vbcrlf
next i

end function
