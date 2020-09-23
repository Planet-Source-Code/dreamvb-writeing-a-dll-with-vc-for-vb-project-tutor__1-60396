// VBTest.cpp : Defines the entry point for the DLL application.

#include "stdafx.h"

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    return TRUE;
}

int _stdcall Add(int x,int y)
{
	//Function to add the values of x and y
	return x+y;
}
