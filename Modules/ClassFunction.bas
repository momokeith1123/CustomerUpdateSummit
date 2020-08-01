Attribute VB_Name = "ClassFunction"
Option Explicit
Function GetContactList(CustId As String, rst As Recordset) As String
    
    'Dim rst                 As ADODB.Recordset
    Dim sql_SELECT          As String
    Dim sql_FROM            As String
    Dim sql_WHERE           As String
    Dim sql_Query           As String
    Dim ContactList         As Dictionary
    Dim newContact          As cContact
    Dim strContact          As String
    Dim i                   As Integer
    
    'sql_Query = LireFichierTexteParLigne("Y:\Users\Mamadou\01-Dev\03-ScriptSQL\Customer\ContactList.sql")
    'sql_Query = sql_Query
    'sql_Query = sql_Query & vbCrLf & "and  cust.id = '" & CustId & "'"
    
    'If InitConnection("ORACLE_TEST", "v56j1", "v56j1") Then
       
    '   Set rst = New ADODB.Recordset
    Set ContactList = New Dictionary
       
    '   If ExecSQL(sql_Query, rst, ADOcnx) Then
            rst.MoveLast
            rst.MoveFirst
            
            strContact = Chr(0)
            With rst
                While (Not .EOF)
                    
                    If rst.Fields(PK) = CustId Then
                        Set newContact = New cContact
                        newContact.CustId = IIf(Len(Trim(rst.Fields("CUSTID"))) > 0, "(CustId:" & """" & Trim(rst.Fields("CUSTID")) & """", "")
                        newContact.ADDR_FullName = IIf(Len(Trim(rst.Fields("ADDR_FULLNAME"))) > 0, " ADDR_FullName:" & """" & Trim(rst.Fields("ADDR_FULLNAME")) & """", "")
                        
                        newContact.ADDR_Addr1 = IIf(Len(Trim(rst.Fields("ADDR_ADDR1"))) > 0, " ADDR_ADDR1:" & """" & Trim(rst.Fields("ADDR_ADDR1")) & """", "")
                        newContact.ADDR_Addr2 = IIf(Len(Trim(rst.Fields("ADDR_ADDR2"))) > 0, " ADDR_ADDR2:" & """" & Trim(rst.Fields("ADDR_ADDR2")) & """", "")
                        newContact.ADDR_Addr3 = IIf(Len(Trim(rst.Fields("ADDR_ADDR3"))) > 0, " ADDR_ADDR3:" & """" & Trim(rst.Fields("ADDR_ADDR3")) & """", "")
                        
                        newContact.ADDR_Country = IIf(Len(Trim(rst.Fields("ADDR_COUNTRY"))) > 0, " ADDR_Country:" & """" & Trim(rst.Fields("ADDR_COUNTRY")) & """", "")
                        
                        newContact.ADDR_ZipCode = IIf(Len(Trim(rst.Fields("ADDR_ZIPCODE"))) > 0, " ADDR_ZIPCODE:" & """" & Trim(rst.Fields("ADDR_ZIPCODE")) & """", "")
                        
                        newContact.Phone = IIf(Len(Trim(rst.Fields("PHONE"))) > 0, " Phone:" & """" & Trim(rst.Fields("PHONE")) & """", "")
                        
                        newContact.fax = IIf(Len(Trim(rst.Fields("Fax"))) > 0, " FAX:" & """" & Trim(rst.Fields("FAX")) & """", "")
                        newContact.telex = IIf(Len(Trim(rst.Fields("TELEX"))) > 0, " Telex:" & """" & Trim(rst.Fields("TELEX")) & """", "")
                        newContact.TELEXANSWERBACK = IIf(Len(Trim(rst.Fields("TelexAnswerBack"))) > 0, " TelexAnswerBack:" & """" & Trim(rst.Fields("TelexAnswerBack")) & """", "")
                        
                        newContact.Department = IIf(Len(Trim(rst.Fields("DEPARTMENT"))) > 0, " Department:" & """" & Trim(rst.Fields("DEPARTMENT")) & """", "")
                        newContact.Email = IIf(Len(Trim(rst.Fields("EMAIL"))) > 0, " Email:" & """" & Trim(rst.Fields("EMAIL")) & """", "")
                        newContact.Comment1 = IIf(Len(Trim(rst.Fields("COMMENT1"))) > 0, " Comment1:" & """" & Trim(rst.Fields("COMMENT1")) & """", "")
                        newContact.Comment2 = IIf(Len(Trim(rst.Fields("COMMENT2"))) > 0, " Comment2:" & """" & Trim(rst.Fields("COMMENT2")) & """", "")
                        
                        newContact.Data = newContact.CustId & newContact.ADDR_FullName & newContact.ADDR_Addr1 & newContact.ADDR_Addr2 & newContact.ADDR_Addr3 & newContact.ADDR_Country _
                        & newContact.ADDR_ZipCode & newContact.Phone & newContact.fax & newContact.telex & newContact.TELEXANSWERBACK & newContact.Department & newContact.Email & newContact.Comment1 & newContact.Comment2 & Chr(32) & ")"
                        
                        If Not ContactList.Exists(newContact.Data) Then
                            ContactList.Add newContact.Data, newContact
                            strContact = Trim(strContact) & Trim(newContact.Data)
                        Else
                            Stop
                        End If
                    Else
                    End If
                    .MoveNext
                Wend
            End With
       'Else
       ' Stop
       'End If
'    Else
'        Stop
'    End If
     If Len(Trim(strContact)) > 3 Then
        strContact = IIf(Left(strContact, 1) = "(", strContact, Right(strContact, Len(strContact) - 1))
        GetContactList = "ContactList:[|CONTACT|" & Chr(32) & strContact & "]"
     Else
        GetContactList = ""
     End If
 End Function

Function GetClassifAndRoleList(CustId As String, rst As Recordset) As String
    'Dim rst                 As ADODB.Recordset
    Dim sql_Query           As String
    Dim DicoList            As Dictionary
    Dim newCust             As cClassif
    Dim strClass            As String
    Dim strData             As String
    Dim i                   As Integer
    
    Dim strRole             As String
    Dim Role                As String
    Dim adb As String, fitch As String, moodys As String, snp As String, role_afdb As String, role_fitch As String, role_moodys As String, role_snp As String
    Dim stype As String
    
       
    'Set rst = New ADODB.Recordset
    'Set DicoList = New Dictionary
    
    
    rst.MoveLast
    rst.MoveFirst
    
    strData = Chr(0)
    strRole = Chr(0)
    
    With rst
        While (Not .EOF)
            If rst.Fields(PK) = CustId Then
                Set newCust = New cClassif
                      
                      
                adb = IIf(IsNull(rst.Fields("ADB_AGCY")), "NR", rst.Fields("ADB_AGCY"))
                fitch = IIf(IsNull(rst.Fields("FITCH")), "NR", rst.Fields("FITCH"))
                moodys = IIf(IsNull(rst.Fields("MOODYS")), "NR", rst.Fields("MOODYS"))
                snp = IIf(IsNull(rst.Fields("SNP")), "NR", rst.Fields("SNP"))
                stype = IIf(IsNull(rst.Fields("TYPE")), "NA", rst.Fields("TYPE"))
                
                role_afdb = IIf(IsNull(rst.Fields("ROLE_AFDB")), "NA", rst.Fields("ROLE_AFDB"))
                role_fitch = IIf(IsNull(rst.Fields("ROLE_FITCH")), "NA", rst.Fields("ROLE_FITCH"))
                role_moodys = IIf(IsNull(rst.Fields("ROLE_MOODYS")), "NA", rst.Fields("ROLE_MOODYS"))
                role_snp = IIf(IsNull(rst.Fields("ROLE_SNP")), "NA", rst.Fields("ROLE_SNP"))
                
                newCust.afdb = getRating("ADB_AGCY", adb, stype, role_afdb)
                newCust.fitch = getRating("FITCH", fitch, stype, role_fitch)
                newCust.moodys = getRating("MOODYS", moodys, stype, role_moodys)
                newCust.snp = getRating("S&P", snp, stype, role_moodys)
                
                newCust.CustRole = "(" & "Custrole:" & """" & rst.Fields("CUSTROLE") & """" & Chr(32) & ")"
                
                strRole = newCust.CustRole & strRole
                strData = newCust.afdb & newCust.fitch & newCust.moodys & newCust.snp
            End If
            .MoveNext
        Wend
        'Stop
    End With
    
    If Len(Trim(strData) > 3) Then
        If Right(strRole, 1) = ")" Then
            Role = strRole & "]"
        Else
            Role = Left(strRole, Len(strRole) - 1) & "]"
        End If
        GetClassifAndRoleList = "ClassifList:[|CUSTCLAS|" & Chr(32) & strData & "]" & Chr(32) & "CustRoleList:[|CUSTROLE|" & Chr(32) & Role
    Else
        GetClassifAndRoleList = ""
    End If
 End Function
 Function GetBKCode(CustId As String, rst As Recordset) As String
    'Dim rst                 As ADODB.Recordset
    Dim sql_Query           As String
    Dim DicoList            As Dictionary
    Dim newCust             As cBkCodeList
    Dim strClass            As String
    Dim strData             As String
    Dim i                   As Integer
    
     rst.MoveLast
     rst.MoveFirst
     
     strData = Chr(0)
     With rst
         While (Not .EOF)
             If rst.Fields(PK) = CustId Then
                Set newCust = New cBkCodeList
                newCust.BkCode = IIf(Len(Trim(rst.Fields("BKCODE"))) > 0, "(BkCode:" & """" & Trim(rst.Fields("BKCODE")) & """" & Chr(32), "")
                newCust.BkValue = IIf(Len(Trim(rst.Fields("BKVALUE"))) > 0, "BkValue:" & """" & Trim(rst.Fields("BKVALUE")) & """" & Chr(32) & ")", "")
                
                newCust.Data = newCust.BkCode & newCust.BkValue
                strData = strData & newCust.Data
               
             Else
             End If
             
             
             .MoveNext
         Wend
     End With
     
     If Len(Trim(strData)) > Len("BkCodeList:[|CUSBKCD|") Then
        If Left(strData, 1) = "(" Then
            
        Else
            strData = Right(strData, Len(strData) - 1)
        End If
         GetBKCode = "BkCodeList:[|CUSBKCD|" & strData & "]"
     Else
         GetBKCode = ""
     End If
     
'     If Len(Trim(strData) > 3) Then
'        If Right(strRole, 1) = ")" Then
'            Role = strRole & "]"
'        Else
'            Role = Left(strRole, Len(strRole) - 1) & "]"
'        End If
'        GetClassifAndRoleList = "ClassifList:[|CUSTCLAS|" & Chr(32) & strData & "]" & Chr(32) & "CustRoleList:[|CUSTROLE|" & Chr(32) & Role
'    Else
'        GetClassifAndRoleList = ""
'    End If
    
 End Function
Function GetMasterAGList(CustId As String, rst As Recordset) As String
    'Dim rst                 As ADODB.Recordset
    Dim sql_Query           As String
    Dim DicoList            As Dictionary
    Dim newCust             As cMasterAGList
    Dim strClass            As String
    Dim strData             As String
    Dim i                   As Integer
    
     rst.MoveLast
     rst.MoveFirst
     
     strData = Chr(0)
     With rst
         While (Not .EOF)
             If rst.Fields(PK) = CustId Then
                Set newCust = New cMasterAGList
                
             Else
             End If
             
             
             .MoveNext
         Wend
     End With
     
     If Len(strData) > 0 Then
         GetMasterAGList = "MasterAGList:[|CUSTMA|" & strData & "]"
     Else
         GetMasterAGList = ""
     End If
    
 End Function

Function GetCustomerId(CustId As String, rst As Recordset) As String
    Dim rstfind                 As ADODB.Recordset
    Dim sql_Query           As String
    Dim DicoList            As Dictionary
    Dim newCust             As cID
    Dim strClass            As String
    Dim strData             As String
    Dim i                   As Integer
    Dim Criterion           As String
    Dim ID_1, ID_2 As String
    
    
    
    Criterion = ""
     rst.MoveLast
     rst.MoveFirst
     
     rst.Find PK & "='" & CustId & "'"
     
     If rst.EOF Then
        '   No records matching search
     Else
     'getData(sfield As String, stext As String, rst As Recordset)
                Set newCust = New cID
                newCust.ID = getData(PK, "Id", rst)
                newCust.INPUTDATE = getData("INPUTDATE", "Inputdate", rst)
                newCust.SHORTNAME = getData("SHORTNAME", "ShortName", rst)
                newCust.LEGALNAME = getData("LEGALNAME", "LegalName", rst)
                newCust.CONTACTNAME = getData("CONTACTNAME", "ContactName", rst)
                newCust.FULLNAME = getData("FULLNAME", "FullName", rst)
                newCust.ADLINE1 = getData("ADLINE1", "AdLine1", rst)
                newCust.ADLINE2 = getData("ADLINE2", "AdLine2", rst)
                newCust.ADLINE3 = getData("ADLINE3", "AdLine3", rst)
                newCust.ADLINE4 = getData("ADLINE4", "AdLine4", rst)
                newCust.ADLINE5 = getData("ADLINE5", "AdLine5", rst)
                newCust.COUNTRY = getData("COUNTRY", "Country", rst)
                newCust.ZIPCODE = getData("ZIPCODE", "ZipCode", rst)
                newCust.FAXNO = getData("FAXNO", "FaxNo", rst)
                newCust.TELEXNUMBER = getData("TELEXNUMBER", "TelexNumber", rst)
                newCust.TELEXANSWERBACK = getData("TELEXANSWERBACK", "TelexAnswerback", rst) '"TELEXANSWERBACK"
                newCust.PHONENUMBER = getData("PHONENUMBER", "PhoneNumber", rst)
                newCust.LOGICALCOUNTRY = getData("LOGICALCOUNTRY", "LogicalCountry", rst)
                newCust.LOGICALFOREIGN = getData("LOGICALFOREIGN", "LogicalForeign", rst)
                newCust.LOGICALCONTINENT = getData("LOGICALCONTINENT", "LogicalContinent", rst)
                newCust.PHYSICALCOUNTRY = getData("PHYSICALCOUNTRY", "PhysicalCountry", rst)
                newCust.PHYSICALFOREIGN = getData("PHYSICALFOREIGN", "PhysicalForeign", rst)
                newCust.PHYSICALCONTINENT = getData("PHYSICALCONTINENT", "PhysicalContinent", rst)
                newCust.SENSITIVE = getData("SENSITIVE", "Sensitive", rst)
                newCust.CONFIRMLANG = getData("CONFIRMLANG", "ConfirmLang", rst)
                newCust.CONFIRMREQ = getData("CONFIRMREQ", "ConfirmReq", rst)
                newCust.BANKINDICATOR = getData("BANKINDICATOR", "BankIndicator", rst)
                newCust.PARENT = getData("PARENT", "Parent", rst)
                newCust.COMPANYAFFIL = getData("COMPANYAFFIL", "CompanyAffil", rst)
                newCust.PARENTAFFIL = getData("PARENTAFFIL", "ParentAffil", rst)
                newCust.LBSFAFFIL = getData("LBSFAFFIL", "LBSFAffil", rst)
                newCust.ISDAAFFIL = getData("ISDAAFFIL", "ISDAAffil", rst)
                newCust.GROUPID = getData("GROUPID", "GroupID", rst)
                newCust.COOKERATIO = getData("COOKERATIO", "CookeRatio", rst)
                newCust.DEFCREDITGROUP = getData("DEFCREDITGROUP", "DEFCREDITGROUP", rst)
                newCust.CUSTTYPE = getData("CUSTTYPE", "CustType", rst)
                newCust.GLACCOUNT = getData("GLACCOUNT", "GLACCOUNT", rst)
                newCust.DEFCARRIER = getData("DEFCARRIER", "DefCarrier", rst)
     End If
    
    
     GetCustomerId = "|CUSTOMER| (" & newCust.ID & newCust.SHORTNAME & Chr(32) & newCust.LEGALNAME & Chr(32) & newCust.CONTACTNAME & Chr(32) & newCust.FULLNAME & Chr(32) _
                  & newCust.ADLINE1 & Chr(32) & newCust.ADLINE2 & Chr(32) & newCust.ADLINE3 & Chr(32) & newCust.ADLINE4 & Chr(32) & newCust.ADLINE5 & Chr(32) _
                  & newCust.COUNTRY & Chr(32) & newCust.ZIPCODE & Chr(32) & newCust.FAXNO & Chr(32) & newCust.TELEXNUMBER & Chr(32) & newCust.TELEXANSWERBACK & Chr(32) _
                  & newCust.PHONENUMBER & Chr(32) & newCust.LOGICALCOUNTRY & Chr(32) & newCust.LOGICALFOREIGN & Chr(32) & newCust.LOGICALCONTINENT & Chr(32) _
                  & newCust.PHYSICALCOUNTRY & Chr(32) & newCust.PHYSICALFOREIGN & Chr(32) & newCust.PHYSICALCONTINENT & Chr(32) & newCust.SENSITIVE & Chr(32) _
                  & newCust.CONFIRMLANG & Chr(32) & newCust.CONFIRMREQ & Chr(32) & newCust.BANKINDICATOR & Chr(32) & newCust.PARENT & Chr(32) & newCust.COMPANYAFFIL & Chr(32) _
                  & newCust.PARENTAFFIL & Chr(32) & newCust.LBSFAFFIL & Chr(32) & newCust.ISDAAFFIL & Chr(32) & newCust.GROUPID & Chr(32) & newCust.COOKERATIO & Chr(32) _
                  & newCust.DEFCREDITGROUP & Chr(32) & newCust.CUSTTYPE & Chr(32) & newCust.GLACCOUNT & Chr(32) & newCust.DEFCARRIER & Chr(32)


                                    
                                    
                        
        
 End Function




