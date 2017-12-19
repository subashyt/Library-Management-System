Attribute VB_Name = "Module1"
Type book
 'to store records details of books
 bookid As String * 12
 bname As String * 50
 author As String * 50
 categ As String * 30
 price As Currency
End Type

Type member
 'to store record details of members
 memid As String * 12
 mname As String * 40
 phone As String * 10
 addre As String * 80
 doj As Date
 valid As Date
 sex As Boolean
End Type

Type issue
 'to manipulate issue records
 ibookid As String * 12
 imemid As String * 12
 idate As Date
End Type

Type login
 'to store login details
 password As String
 'oldpass As String
 End Type

