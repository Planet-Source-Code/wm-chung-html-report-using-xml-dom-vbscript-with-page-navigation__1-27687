		sub nextpage()
		if (xmldso.Recordset.absoluteposition < xmldso.Recordset.recordcount)  then
		xmldso.recordset.movenext
		end if
		WritePageNo
		end sub


		sub PreviousPage()
		if xmldso.Recordset.absoluteposition <> 1 then
		xmldso.recordset.moveprevious
		end if
		WritePageNo
		end sub



		sub LastPage()
		xmldso.recordset.movelast
		WritePageNo
		displaysummary
		end sub
		
		
		sub WritePageNo()
		pagefield.innerHTML="Page " & cstr(xmldso.recordset.absoluteposition)
		end sub

sub WriteDate()
rundate.innerHTML=datevalue(now())'cstr(format(now, "ddmmyy"))

end sub



sub writeheadingxp()
xmldso.src = "xmlexp" & day(now()) & month(now()) & year(now()) & ".xml"
writepageno()
writedate()

end sub




sub writeheading()

writepageno()
writedate()
end sub



sub writeheadingbudget()
'rpttitle.innerhtml = "<center>REPORT FOR THE MONTH " & month(now) & "/" & year(now) & "</center>"
writepageno()
writedate()
end sub

		sub FirstPage()
		
		xmldso.recordset.movefirst
		WritePageNo

		end sub
		
		
sub cleartable()

		cap_overrun.innerhtml = ""
		cap_ytdbudget.innerhtml = ""
			cap_ytdamt.innerhtml = ""
		cap_curbudget.innerhtml = ""
			cap_curexp.innerhtml = ""
					rec_overrun.innerhtml = ""
		rec_ytdbudget.innerhtml = ""
			rec_ytdamt.innerhtml = ""
		rec_curbudget.innerhtml = ""
			rec_curexp.innerhtml = ""
			rec_descript.innerhtml = ""
			cap_descript.innerhtml = ""
end sub		


sub displaybudgetsummary()
	dim ytdbudget, ytdactual, currentbudget, currentexp
	dim capitalitem, recurrentitem
	dim strout
	set capitalitem = new budgetsummary
	set recurrentitem = new budgetsummary

	Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
	objXMLDoc.async = False 
	objXMLDoc.load("xmlbudget.xml") 

	Dim root, i, j, k, targetnode
	Set root = objXMLDoc.documentElement 
		For i = 0 To (root.childNodes.length-1)

			for  subk =0 to root.childnodes.item(i).childnodes.length-1

				for subl = 0 to root.childnodes.item(i).childnodes.item(subk).childnodes.length -1
				set targetnode = root.childnodes.item(i).childnodes.item(subk).childnodes'.item(subl)
				'document.write(targetnode.nodename & " | ")
					if strcomp(targetnode.item(subl).nodename, "account_no") = 0 then
						if strcomp(mid(targetnode.item(subl).text, 1,1), "1") = 0 then
							capitalitem.ytdamount  = capitalitem.ytdamount +int(targetnode.item(subl+1).text)
							capitalitem.ytdbudget = capitalitem.ytdbudget  +int(targetnode.item(subl+3).text)
							capitalitem.currentbudget  = capitalitem.currentbudget  +int(targetnode.item(subl+4).text)
							capitalitem.currentexp = capitalitem.currentexp   +int(targetnode.item(subl+5).text)
							capitalitem.overrun = capitalitem.overrun  +int(replace(targetnode.item(subl+6).text, "*", ""))
							
						elseif strcomp(mid(targetnode.item(subl).text, 1,1), "2") = 0 then
							recurrentitem.ytdamount  = recurrentitem.ytdamount +int(targetnode.item(subl+1).text)
							recurrentitem.ytdbudget = recurrentitem.ytdbudget  +int(targetnode.item(subl+3).text)
							recurrentitem.currentbudget  = recurrentitem.currentbudget  +int(targetnode.item(subl+4).text)
							recurrentitem.currentexp = recurrentitem.currentexp   +int(targetnode.item(subl+5).text)
							recurrentitem.overrun = recurrentitem.overrun  +int(replace(targetnode.item(subl+6).text, "*", ""))
						end if
					end if
				next		
			next
		Next
		cap_descript.innerhtml = "Capital Expenditure Details :"
		cap_overrun.innerhtml = "<p align = 'right'>" & capitalitem.overrun & "</P>"
		cap_ytdbudget.innerhtml = "<p align = 'right'>" & capitalitem.ytdbudget & "</P>"
			cap_ytdamt.innerhtml = "<p align = 'right'>" & capitalitem.ytdamount  & "</P>"
		cap_curbudget.innerhtml = "<p align = 'right'>" & capitalitem.currentbudget & "</P>"
			cap_curexp.innerhtml = "<p align = 'right'>" & capitalitem.ytdamount & "</P>"
				rec_descript.innerhtml = "Recurrent Expenditure Details :"
					rec_overrun.innerhtml = "<p align = 'right'>" & recurrentitem.overrun & "</P>"
		rec_ytdbudget.innerhtml = "<p align = 'right'>" & recurrentitem.ytdbudget & "</P>"
			rec_ytdamt.innerhtml = "<p align = 'right'>" & recurrentitem.ytdamount  & "</P>"
		rec_curbudget.innerhtml = "<p align = 'right'>" & recurrentitem.currentbudget & "</P>"
			rec_curexp.innerhtml = "<p align = 'right'>" & recurrentitem.ytdamount & "</P>"
		
end sub




Class budgetsummary 

   ' Creating a private property using Get, Let, Set 
   Private ybudget, yamt, cbudget, cexp, over 
   ' Get 
   Public Property Get ytdbudget() 
      ytdbudget = ybudget 
   End Property 
   ' Let 
   Public Property Let ytdbudget(strName) 
      ybudget = strName 
   End Property 
      Public Property Get ytdamount() 
      ytdamount = yamt 
   End Property 
   ' Let 
   Public Property Let ytdamount(strName) 
      yamt = strName 
   End Property 
      Public Property Get currentbudget() 
      currentbudget = cbudget 
   End Property 
   ' Let 
   Public Property Let currentbudget(strName) 
      cbudget = strName 
   End Property 
      Public Property Get currentexp() 
      currentexp = cexp 
   End Property 
   ' Let 
   Public Property Let currentexp(strName) 
      cexp = strName 
   End Property 
      Public Property Get overrun() 
      overrun = over 
   End Property 
   ' Let 
   Public Property Let overrun(strName) 
      over = strName
   End Property 


End Class 



sub ShowHideSummary()
	if len(summary.innerhtml) > 0 then
		summary.innerhtml = ""
		showhide.value = "Display Summary"

	else
		DisplaySummary()
	showhide.value = "Hide Summary"
	end if

end sub