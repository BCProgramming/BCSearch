Attribute VB_Name = "ModUnitTest"
Option Explicit

'UnitTests for BCFile.



Sub Main()

    Con.Initialize
    Con.WriteLine "BCFile Unit Test Module."
    'create a test stream.
    Dim fs As FileStream
    
    Set fs = BCFile.CreateStream("D:\testoutput.txt")
    Dim I As Long
    For I = 0 To 100
    fs.WriteLine "Greetings, this is a test file, line #" & I
    Next I
    fs.CloseStream
    
    Con.WriteLine ("Reading...")
    Set fs = BCFile.OpenStream("D:\testoutput.txt")
    Do While Not fs.AtEndOfStream
    Dim readline As String
    readline = fs.readline()
    Con.WriteLine readline
    
    Loop
    
    BCFile.ShellExec 0, "D:\testoutput.txt"
    


End Sub
