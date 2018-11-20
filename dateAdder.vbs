dim dateIn
dim dateOut
dim nDays

dateIn = WScript.Arguments.Item(0)
nDays = WScript.Arguments.Item(1)

dateOut = DateAdd("d", nDays, dateIn)

WScript.StdOut.WriteLine dateOut