Set WshShell= WScript.Createobject("WScript.Shell")
　　for i=1 to 9999
　　WScript.Sleep 700
　　WshShell.SendKeys"^v"
　　WshShell.SendKeys i
　　WshShell.SendKeys "%s"
　　next