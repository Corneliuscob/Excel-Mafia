# Mafia written in VBA


###The following code is written in VBA to simulate 1 night of the game Mafia. It uses Collections as the way to manage all the actions that occur and use the Excel Spreadsheet as a "database".


### amain is the main module and  function which performs the nights actions that occur. First, it creates the player colelction which is the current state of the Excel Sheet. This code is stored in the Creat_collection module.


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

### Each player action follows a similar  structure. The majority of the player methods take a target, themselves(the character/activeUser), the player collection, as well as a depth counter and a former target (last two being optional parameters. They all have this syntax

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


