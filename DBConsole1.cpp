// DBConsole1.cpp : This file contains the 'main' function. Program execution begins and ends there.
//
#include "stdafx.h"
#include <iostream>
#include "windows.h"
#import "C:\Program Files\Common Files\System\ado\msado60_Backcompat.tlb" no_namespace rename("EOF", "EOFile")

int main()
{
    // define our variables which will be used as references to the
    // Connection and Recordset objects
    _ConnectionPtr con = NULL;
    _RecordsetPtr rec = NULL;
    // define variables to read the Author field from the recordset
    FieldPtr pAuthor;
    _variant_t vAuthor;
    char sAuthor[40];
    // create two strings for use with the creation of a Connection
    // and a Recordset object
    bstr_t sConString;
    bstr_t sSQLString;
    // create a variable to hold the result to function calls
    HRESULT hr = S_OK;
    // long variable needed for Execute method of Connection object
    VARIANT* vRecordsAffected = NULL;


    if (FAILED(hr = CoInitialize(NULL)))
    {
        return 0;
    }
    try
    {
        // create a new instance of an ADO Connection object
        hr = con.CreateInstance(__uuidof(Connection));
        printf("Connection object created.\n");
        con->put_CursorLocation(adUseClient);
        // open the BiblioDSN data source with the Connection object

        /*sConString = L"Provider=.NET Framework Data Provider for SQL Server; Data Source=DESKTOP-NUKCG05;Initial Catalog=AdventureWorks2017;Integrated Security=True";*/
        /*"DRIVER = { SQL Server }; SERVER = localhost, 1433; DATABASE = AdventureWorks2017; Integrated Security = SSPI; "*/
        sConString = "DRIVER = { SQL Server }; SERVER = localhost, 1433; DATABASE = AdventureWorks2017; Integrated Security = SSPI; ";
        con->Provider = "sqloledb";
        con->Open(sConString, "", "", adConnectUnspecified);
        printf("Connection opened.\n");
        // create a Recordset object from a SQL string
        sSQLString = L"select Name from Production.Location;";
        rec = con->Execute(sSQLString, vRecordsAffected, 1);
        printf("SQL statement processed.\n");
        // point to the Author field in the recordset
        pAuthor = rec->Fields->GetItem("Name");
        // retrieve all the data within the Recordset object
        printf("Getting data now...\n\n");
        while (!rec->EOFile)
        {
            //std::cout << "Hello World!\n";
            // get the Author field's value and change it
            // to a multibyte type
            vAuthor.Clear();
            vAuthor = pAuthor->Value;
            WideCharToMultiByte(CP_ACP,
                0,
                vAuthor.bstrVal,
                -1,
                sAuthor,
                sizeof(sAuthor),
                NULL,
                NULL);
            printf("%s\n", sAuthor);
            rec->MoveNext();
        }
        printf("\nEnd of data.\n");
        // close and remove the Recordset object from memory
        rec->Close();
        rec = NULL;
        printf("Closed an removed the "
            "Recordset object from memory.\n");
        // close and remove the Connection object from memory
        con->Close();
        con = NULL;
        printf("Closed and removed the "
            "Connection object from memory.\n");
    }
    catch (std::exception e)
    {
        printf(e.what());
    }
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file



// Visual_CPP_ADO_Prog_1.cpp  
// compile with: /EHsc  
//#import "msado15.dll" no_namespace rename("EOF", "EndOfFile")  

// Note 1  
//inline void TESTHR(HRESULT _hr) {
//    if FAILED(_hr)
//        _com_issue_error(_hr);
//}
//
//int main() {
//    CoInitialize(NULL);
//    try {
//        _RecordsetPtr pRs("ADODB.Recordset");
//        _ConnectionPtr pCn("ADODB.Connection");
//        _variant_t vtTableName("authors"), vtCriteria;
//        long ix[1];
//        SAFEARRAY* pSa = NULL;
//
//        pCn->Provider = "sqloledb";
//        //"Data Source='(local)';Initial Catalog='pubs';Integrated Security=SSPI"
//        pCn->Open("DRIVER = { SQL Server }; SERVER = localhost, 1433; DATABASE = AdventureWorks2017; Integrated Security = SSPI; ", "", "", adConnectUnspecified);
//        // Note 2, Note 3  
//        pSa = SafeArrayCreateVector(VT_VARIANT, 1, 4);
//        if (!pSa)
//            _com_issue_error(E_OUTOFMEMORY);
//
//        // Specify TABLE_NAME in the third array element (index of 2).   
//        ix[0] = 2;
//        TESTHR(SafeArrayPutElement(pSa, ix, &vtTableName));
//
//        // There is no Variant constructor for a SafeArray, so manually set the   
//        // type (SafeArray of Variant) and value (pointer to a SafeArray).  
//
//        vtCriteria.vt = VT_ARRAY | VT_VARIANT;
//        vtCriteria.parray = pSa;
//
//        pRs = pCn->OpenSchema(adSchemaColumns, vtCriteria, vtMissing);
//
//        long limit = pRs->GetFields()->Count;
//        for (long x = 0; x < limit; x++)
//            printf("%d: %s\n", x + 1, ((char*)pRs->GetFields()->Item[x]->Name));
//        // Note 4  
//        pRs->Close();
//        pCn->Close();
//    }
//    catch (_com_error& e) {
//        printf("Error:\n");
//        printf("Code = %08lx\n", e.Error());
//        printf("Code meaning = %s\n", (char*)e.ErrorMessage());
//        printf("Source = %s\n", (char*)e.Source());
//        printf("Description = %s\n", (char*)e.Description());
//    }
//    CoUninitialize();
//}