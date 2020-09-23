// vbDll.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "string.h"


BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    return TRUE;
}


int _stdcall ReportVersion()
{
	// report version of this dll only ment as a test
	// between the C++ and VB
	return 1;
}


int _stdcall Add(int x,int y)
{
	// add two numbers
	return x+y;
}

int _stdcall Sum(int *slots,int nSize)
{
	// takes in an array of integer numbers and then sums up the total result
	int x;
	int total=0;
	
	for(x=0;x<nSize;x++)
	{
		total += slots[x];
	}
	
	return total;
}

void _stdcall UcaseEx(char *s)
{
	// upper case a string
	int x;
	for(x=0;x<strlen(s);x++)
	{
		s[x] = toupper(s[x]);
	}
}


char * _stdcall LCaseEx(char *s)
{
	//lower case text
	_strlwr(s);
	return s;
}


int SHL(int x,int places)
{
	//bitshift left
	return x << places;
}

int SHR(int x,int places)
{
	// bit shift right
	return x >> places;
}