Dim Mesaj, Speak
Mesaj=Inputbox("KONU�AN VBS KODU","TORBAC� HUSEY�N")
Set Speak=CreateObject("sapi.spvoice")
Speak.Speak Mesaj
