# Mafia written in VBA


### The following code is written in VBA to simulate 1 night of the game Mafia. It uses Collections as the way to manage all the actions that occur and use the Excel Spreadsheet as a "database".


### amain is the main module and  function which performs the nights actions that occur. First, it creates the player colelction which is the current state of the Excel Sheet. This code is stored in the Creat_collection module.
        Function create_collection()
    ' This function creates a collection of players in from the excel sheet
    ' this function returns the created collection
    
    'dimensions
        Dim CharacterList As Collection
        Dim character As cCharacter
        Set CharacterList = New Collection
    'last row of users
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    'assignment of each character's properties
        For i = 2 To lastrow
            Set character = New cCharacter
                character.name = Cells(i, 1)
                character.Char = Cells(i, 2)
                character.Status = Cells(i, 3)
                character.target = Cells(i, 4)
                character.Affiliation = Cells(i, 5)
                character.Special = Cells(i, 6)
                character.Doused = Cells(i, 7)
                
                character.Healed = Cells(i, 8)
                character.Protected = Cells(i, 9)
                character.Blocked = Cells(i, 10)
                character.Visit = Cells(i, 11)
            CharacterList.Add character
        Next i
        Set create_collection = CharacterList
	End Function

### Then, cycle actions occur. This is a wrapper loops through the collection for the players who  are character X ( mafioso, doctors,  bodyguards), following the game mechanics and performs the actions that the characters would have. For example a mafioso would have the ability to shoot a target. The cycle actions are found in the cycle_actions module. They all follow the syntax.


        Sub cycleArsonists(coll As Collection) 
                Dim user As cCharacter `
                For Each user In coll
                    If user.Char = "arsonist" Then
                        If user.target = "ignite" Then
                            ArsonistIgnite user.name, coll
                        ElseIf user.target = "" Then
                        Else
                            ArsonistDouse user.target, user.name, coll
                        End If
                    End If
                Next user
         End Sub  `

### Each player action follows a similar structure. The majority of the player methods take a target, themselves(the character/activeUser), the player collection, as well as a depth counter and a former target (last two being optional parameters. They all have this syntax: The target and active user index are found using a function. The active user is checked to see if they were blocked. The person the active user visits is addressed (in case the active user is redirected by a busdriver). They check that the target is not the veteran  on alert. They check whether or not hte user has been switched. If the user is switched the function is recursively called once (using the depth += 1) as the former target is still the place visited. After all the checks, the character specific action happens. In the code below the investigate action happens. The possible crimes are displayed as a string. Up until the line with `result` the code is the samefor most player actions, just written because vba does not support inheritance (to my knoowledge). 

        Sub Investigate(target As String, activeUser As String, ccoll As Collection, Optional depth As Integer = 1, Optional former As String)
            Dim uList As String
            ''this sub takes the target of the activeUser,  the activeUser and the player collection
            ' this sub prints the target of the activeUser's target
            'this sub is dependent on find_person_index and set_Excel

            target_Id = find_person_index(target, ccoll)
            activeUser_Id = find_person_index(activeUser, ccoll)

            If personIsBlocked(activeUser_Id, ccoll) Then
                setExcel activeUser, 12, "you were blocked"
                Exit Sub
            End If

            If depth > 1 Then
                setExcel activeUser, 11, former
                ccoll.Item(activeUser_Id).Visit = former
            Else
                setExcel activeUser, 11, target
                ccoll.Item(activeUser_Id).Visit = target
            End If

            If targetIsOnAlert(target_Id, ccoll) Then
                ccoll.Item(activeUser_Id).Status = "dead"
                setExcel activeUser, 3, "dead"
                setExcel activeUser, 12, "you were killed by the veteran"
                Exit Sub
            End If

            If ccoll.Item(target_Id).Special <> "" And depth = 1 Then
                Investigate ccoll.Item(target_Id).Special, activeUser, ccoll, depth + 1, target
                Exit Sub
            End If

            result = possible_crime_committed(ccoll.Item(target_Id).Char)

            setExcel activeUser, 12, "your target committed the following crime: " & result


        End Sub

### helper functions are employed to be used in support of the player actions. possible\_crimes\_committed, above is an example of a helper function that is unique to the player.
		
		Function possible_crime_committed(character As String)
    
    
			    		Select Case character
			        Case "villager", "doctor", "sheriff", "jester", "witch", "amnesiac"
			            possible_crime_committed = "no crime"
			        Case "mayor", "marshall"
			            possible_crime_committed = "corruption"
			        Case "investigator", "vigilante", "detective", "lookout", "mafioso", "godfather", "consigliere", "framer", "janitor", "beguiler", "serial_killer", "arsonist", "mass_murderer"
			            possible_crime_committed = "trespassing"
			        Case "consort", "escort"
			            possible_crime_committed = "soliciting"
			        Case "veteran", "janitor", "arsonist", "mass_muderer"
			            possible_crime_committed = "Destruction of property"
    End Select
	End Function

### Two global helper functions are the *find\_person\_index* and *setExcel* function. The first function returns the index at where the  player of interest is stored in the player collection. Once the player is found, the properties of the player can be accessed.

	Function find_person_index(name As String, coll As Collection)
    'this sub finds the number of the person of interest in the PLayers collection
        n = 1
        Dim user As cCharacter
        For Each user In coll
            If user.name = name Then
                find_person_index = n
                'MsgBox "User Found"
                Exit Function
            End If
            n = n + 1
        Next user
	End Function

### The other function is allows us to write back to the Excel sheet the changes that occur. This functino may not be used at a later time and have the collection parsed after the night actions occur. 

	Sub setExcel(target As String, param As Integer, newVal As String, Optional param2 As Integer, Optional newVal2 As String)
	    'This sub takes the player and their parameters of interest and updates the excel sheet
	    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
	    i = 1
	    Do While i <= lastrow
	       If Cells(i, 1).Value = target Then
	            Cells(i, param).Value = newVal
	            If param2 <> 0 Then
	                Cells(i, param2).Value = newVal2
	            End If
	        Exit Do
	       End If
	       i = i + 1
	    Loop
	End Sub
