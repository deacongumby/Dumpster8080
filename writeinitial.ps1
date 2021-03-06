$htmlfile = ([string](get-location) + "\index.html")

if (Test-Path -path $htmlfile) { 
  remove-item $htmlfile
}

Add-Content $htmlfile "<HTML>"
Add-Content $htmlfile ("<TITLE> Status as of " + (get-date) + " </TITLE>")

Add-Content $htmlfile "<style type=`"text/css`">"
Add-Content $htmlfile "td { font-family: times new roman, times, serif; font-size: 13px; }"
Add-Content $htmlfile "</style>"

Add-Content $htmlfile "<img src=`"directoryservices.jpg`"></BR></BR>"

Add-Content $htmlfile "<!--#include file=`"corp.html`"-->"
Add-Content $htmlfile "<!--#include file=`"extranet.html`"-->"
Add-Content $htmlfile "<!--#include file=`"extranettest.html`"-->"
Add-Content $htmlfile "<!--#include file=`"others.html`"-->"

Add-Content $htmlfile "</HTML>"
