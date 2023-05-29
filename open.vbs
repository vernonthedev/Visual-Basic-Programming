Option Explicit

dim ie

set ie = CreateObject("InternetExplorer.Application")
ie.Navigate "https://www.google.com"
ie.Visible = 1
ie.Toolbar = 0