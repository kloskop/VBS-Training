option explicit

Dim MyText, a, wordNum, word
MyText = "The quick brown fox jumps over the lazy dog"
wordNum=0 

For a=1 To Len(MyText)
    If Mid(MyText,a,1)=" " Then 'the logic behind this, is that every word is separated with " " sign, so we can count all of those
		wordNum = wordNum+1
	End If
Next 
MsgBox "There are "&wordNum+1&" words in the sentence" 'adding 1, beacause the last word does not end with " " sign
MsgBox "I found the word: "&mid(MyText,21,5)&" !"
word=mid(MyText,5,5)
MsgBox "I reversed the word: "&word&" to: "&StrReverse(word)+" !"
MsgBox Replace(MyText," ","")
MsgBox "There are "&len(MyText)&" characters in the sentence."
