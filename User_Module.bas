Attribute VB_Name = "Module1"
Type users
id_number As Integer
username As String * 15
password As String * 15
side As Integer
points As Integer
Commander As Integer
ah As Integer
as As Integer
aa As Integer
al As Integer
wins As Integer
losses As Integer
End Type

Type hiscoretable
username As String * 15
turns As Integer
cheated As String * 3
End Type

Type maps
picture As String * 150
loc1left As Integer
loc1top As Integer
loc2left As Integer
loc2top As Integer
loc3left As Integer
loc3top As Integer
loc4left As Integer
loc4top As Integer
loc5left As Integer
loc5top As Integer
loc6left As Integer
loc6top As Integer
loc7left As Integer
loc7top As Integer
loc8left As Integer
loc8top As Integer
loc9left As Integer
loc9top As Integer
loc10left As Integer
loc10top As Integer
loc11left As Integer
loc11top As Integer
loc12left As Integer
loc12top As Integer
loc13left As Integer
loc13top As Integer
loc14left As Integer
loc14top As Integer
End Type

Type listmaps
mapname As String * 20
mappath As String * 150
End Type

Type save
savename As String * 20
mapname As String * 150
loc1 As Integer
loc2 As Integer
loc3 As Integer
loc4 As Integer
loc5 As Integer
loc6 As Integer
loc7 As Integer
loc8 As Integer
loc9 As Integer
loc10 As Integer
loc11 As Integer
loc12 As Integer
loc13 As Integer
loc14 As Integer
usermoney As Integer
userarea As Integer
compmoney As Integer
comparea As Integer
utroops As Integer
uheavy As Integer
utanks As Integer
uspecial1 As Integer
uspecial2 As Integer
ctroops As Integer
cheavy As Integer
ctanks As Integer
cspecial1 As Integer
cspecial2 As Integer
End Type
