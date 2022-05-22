Dim Mesaj, Speak
Mesaj=Inputbox("KONUÞAN VBS KODU","TORBACÝ HUSEYÝN")
Set Speak=CreateObject("sapi.spvoice")
Speak.Speak Mesaj
