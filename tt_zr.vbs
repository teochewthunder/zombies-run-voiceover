Dim speech
Dim item, item_name, quantity, unit, no_items
Dim collections(5), used(5)
Dim items(15)
items(0) = Array("spade", "a", "", "", "spades")
items(1) = Array("toolbox", "a", "", "", "toolboxes")
items(2) = Array("sports bra", "a", "", "", "sports bras")
items(3) = Array("shotgun", "a", "", "", "shotguns")
items(4) = Array("overcoat", "an", "", "", "overcoats")
items(5) = Array("lightbulbs", "", "box", "a", "boxes")
items(6) = Array("money", "", "handful", "a", "handfuls")
items(7) = Array("axe", "an", "", "", "axes")
items(8) = Array("trousers", "", "pair", "a", "pairs")
items(9) = Array("food", "", "tin", "a", "tins")
items(10) = Array("book", "a", "", "", "books")
items(11) = Array("fuel", "", "can", "a", "cans")
items(12) = Array("pain meds", "", "bottle", "a", "bottles")
items(13) = Array("football", "a", "", "", "footballs")
items(14) = Array("radio", "a", "", "", "radios")

Function getrandomno(min, max)
	Randomize
	getrandomno = (Int((max - min + 1) * Rnd + min))
End Function

Function inuse(num)
	Dim dupe

	dupe = false

	for j = 0 to no_items step 1
		if used(j) = num then
	    		dupe = true
		end if
	next

	inuse = dupe
End Function

no_items = getrandomno(1, 4)

for i = 0 to no_items step 1
	do while inuse(item)=true
		item = getrandomno(0, 14)
     	loop

	used(i) = item

	quantity = getrandomno(1, 5)

	unit = ""

	if quantity = 1 then
		if items(item)(2) <> "" then
			unit = " " & items(item)(2) & " of"
			quantity = items(item)(3)
		else
			quantity = items(item)(1)
		end if	

		item_name = items(item)(0)
	else
		if items(item)(2) <> "" then
			unit = " " & items(item)(4) & " of"
			item_name = items(item)(0)
		else
			item_name = items(item)(4)
		end if
	end if

	item_name = " " & quantity & " " & unit & " " & item_name

	collections(i) = item_name
next

'speech = "Collected a"
speech = "Collected"

for i = 0 to no_items step 1
     if i < no_items then
     	speech = speech & collections(i) & ","
     else
	if i = 0 then
		speech = speech & collections(i)
	else
		speech = speech & " and" & collections(i)
	end if
     end if
next

'speech = speech & item_name

Set VObj = CreateObject("sapi.spvoice")
VObj.Speak speech