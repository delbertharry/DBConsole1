// DBConsole1.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include "windows.h"
#import <msado15.dll> no_namespace rename("EOF", "EOFile")

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
    // create a new instance of an ADO Connection object
    hr = con.CreateInstance(_uuidof(Connection));
    printf("Connection object created.\n");
    // open the BiblioDSN data source with the Connection object

    /*sConString = L"Provider=.NET Framework Data Provider for SQL Server; Data Source=DESKTOP-NUKCG05;Initial Catalog=AdventureWorks2017;Integrated Security=True";*/
    sConString = L"Data Source=DESKTOP-NUKCG05;Initial Catalog=AdventureWorks2017;Integrated Security=True"; 
    con->Open(sConString, L"", L"", -1);
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

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
