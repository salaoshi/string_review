// BABYGRID_DEMO.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "resource.h"
#include "babygrid.h"  // <------- You must include the babygrid.h header file
#include "archiving.h"
#include <SHELLAPI.H>
#include <commdlg.h> //dialog

extern int TableRowCount;
int ver=110;
char subver[8]="debug";
#define XLSX 1
#define XML 2

#define MAX_LOADSTRING 100
char ProgramPatch[MAX_PATH];
char Excel_file[MAX_PATH]="";
char Screenshot_col_name[32]="Screenshot";
int Opened_file=0;
/*
struct table
{
int row;
char name[80];
char pic[80];
};*/

struct table *pT;

HWND hWnd;
HWND hScreenshotWnd;

// Global Variables:
HINSTANCE hInst;								// current instance
TCHAR szTitle[MAX_LOADSTRING];					// The title bar text
TCHAR szWindowClass[MAX_LOADSTRING];			// The title bar text

_BGCELL cell;         // <----------- You'll need to define at least one of these
                      //              to reference the grid cells for entering or
                      //              retreiving data from the grid 
HWND hgrid1,hgrid2;   // <------------ Window handles of the grids you'll create

// Foward declarations of functions included in this code module:
ATOM				MyRegisterClass(HINSTANCE hInstance);
BOOL				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK	About(HWND, UINT, WPARAM, LPARAM);

void LoadGrid1(HWND);
void LoadGrid2(HWND);
void SetCell(HWND,int,int,char*);
int GetGurRow();


extern int Table_Target_column ;
extern char Pictures_path[MAX_PATH];
extern unsigned long __stdcall Sender2(void *Param);
void LoadExcelPicture(char* file);
char UserName[28]="";
DWORD User_name_len=sizeof(UserName);
HWND CreatePreviewWindow(HWND ParentHwnd, int x, int y, int dx, int dy);
static DWORD dwId;
HANDLE  hThread_Sender2;


void Word_Count()
{
	char str[1024]="";
								int word_count=0;
								for (int i=2;i<TableRowCount;i++)
								{
									memset(str,0,1024);
									GetCell( hgrid2, i,6, str);

										if(str[0])
										for (int w=0;w<1024;w++)
										{
											if(str[w]==' ') word_count++;
											if(str[w]==0xA) word_count++;
											if(str[w]==0) {word_count++; break;}
										}

								}
								int word_count2=0;
								for (  i=2;i<TableRowCount;i++)
								{
									memset(str,0,1024);
									GetCell( hgrid2, i,7, str);

										if(str[0])
										for (int w=0;w<1024;w++)
										{
											if(str[w]==' ') word_count2++;
											if(str[w]==0) {word_count2++; break;}
										}
								}
			 

								wsprintf(str,"ScreenshotViewer: English words=%d  Edited English words=%d",word_count,word_count2);
								SetWindowText(hWnd,str);
}


void DrawPicture()
{
	char pic_file[64]="";
	int currow=GetGurRow();
	if(Table_Target_column!=-1)
	if(currow>=1)
	{

		GetCell(hgrid2,currow,Table_Target_column,pic_file);

		if(!memcmp(pic_file,"file:",5))
			for(unsigned int i=0;i<strlen(pic_file);i++)
				pic_file[i]=pic_file[i+5];
		
		  LoadExcelPicture( pic_file);
		//PutCell(hgrid2,0,Table_Target_column,"QQ");//
	}
}

int APIENTRY WinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPSTR     lpCmdLine,
                     int       nCmdShow)
{
 	// TODO: Place code here.
	MSG msg;
	HACCEL hAccelTable;

	// Initialize global strings
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_BABYGRID_DEMO, szWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow)) 
	{
		return FALSE;
	}

	hAccelTable = LoadAccelerators(hInstance, (LPCTSTR)IDC_BABYGRID_DEMO);

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0)) 
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg)) 
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	return msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
//  COMMENTS:
//
//    This function and its usage is only necessary if you want this code
//    to be compatible with Win32 systems prior to the 'RegisterClassEx'
//    function that was added to Windows 95. It is important to call this function
//    so that the application will get 'well formed' small icons associated
//    with it.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX); 

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= (WNDPROC)WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, (LPCTSTR)IDI_BABYGRID_DEMO);
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= (LPCSTR)IDC_BABYGRID_DEMO;
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, (LPCTSTR)IDI_SMALL);

	return RegisterClassEx(&wcex);
}

//
//   FUNCTION: InitInstance(HANDLE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//


BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   

   hInst = hInstance; // Store instance handle in our global variable

   hWnd = CreateWindow(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
     CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);
    // 0, 0, 1124,700, NULL, NULL, hInstance, NULL);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, unsigned, WORD, LONG)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND	- process the application menu
//  WM_PAINT	- Paint the main window
//  WM_DESTROY	- post a quit message and return
//
//
int pic_windows_widht=0;
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	PAINTSTRUCT ps;
	HDC hdc;
	TCHAR szHello[MAX_LOADSTRING];
	LoadString(hInst, IDS_HELLO, szHello, MAX_LOADSTRING);

	switch (message) 
	{
		case WM_COMMAND:
			wmId    = LOWORD(wParam); 
			wmEvent = HIWORD(wParam); 
			// Parse the menu selections:
			switch (wmId)
			{
				case IDM_ABOUT:
				   DialogBox(hInst, (LPCTSTR)IDD_ABOUTBOX, hWnd, (DLGPROC)About);
				   break;
				case IDM_EXIT:
				   DestroyWindow(hWnd);
				   break;
				case IDM_OPEN:
					{
					 	char * pFileName;
						OPENFILENAME ofn; 
							ZeroMemory(&ofn, sizeof(OPENFILENAME));
							ZeroMemory(Excel_file, MAX_PATH);
								
									// inizializate OPENFILENAME 
									//ofn.nFilterIndex  =Prev_opened_type;
									ofn.lStructSize = sizeof(OPENFILENAME); 
									ofn.hwndOwner = hWnd; 
									ofn.lpstrFile =Excel_file; 
									ofn.nMaxFile = MAX_PATH; 
									ofn.lpstrFilter =      & "Excel Spreadsheet Files(*.xlsx)\0*.xlsx;*.XLSX\0"  \
															 "Android XML Files(*.xml)\0*.xml;*.XML\0"  \
															"All Files (*.*)\0*.*\0";
															
									//ofn.nFilterIndex = 0; 
									ofn.lpstrTitle=&"Open Excel File"; 
									ofn.lpstrInitialDir=""; 
									ofn.Flags = OFN_EXPLORER|OFN_FILEMUSTEXIST|OFN_ENABLESIZING | OFN_HIDEREADONLY;
									//ofn.lpstrFileTitle=source_file;	
									//ofn.lpTemplateName=source_file; 
									
										//ofn.lpstrFileTitle    = szFileTitle;//get filename witoput path
										//ofn.nMaxFileTitle     = sizeof(szFileTitle);
										
									if(!GetOpenFileName(&ofn)) 
											return 0;

									
									if(Excel_file[0])
									{
										strcpy(Pictures_path,Excel_file);
										pFileName=strrchr(Pictures_path,'\\');
										pFileName++;
										*pFileName=0;					
									}
									if(ofn.nFilterIndex==1)
									{
										Open_Excel_XLSX_file(Excel_file); Opened_file=XLSX ;
										char szSTR[64];
										for(int h=0;h<MAX_COLUMN;h++)
										{
											memset(szSTR,0,64);
											GetCell( hgrid2, 2,h, szSTR);
											if(!memcmp(szSTR,"file:",strlen("file:")))
												Table_Target_column=h;
										}
										Word_Count();
									}


									if(ofn.nFilterIndex==2)
									{
										Open_XMLfile2 (Excel_file); Opened_file=XML ;
										pFileName--;
										*pFileName=0;
										pFileName=strrchr(Pictures_path,'\\');
										pFileName++;
										*pFileName=0;
										Table_Target_column=3;
										  Load_LinkTable(Excel_file);
									}
					}
					break;

				case IDM_SAVE:
					if(Opened_file==XLSX) Save_Excel_XLSX_file(Excel_file);
					if(Opened_file==XML) Save_LinkTable(Excel_file);
					break;

				case IDM_SAVE2:
					if(Opened_file==XLSX) Save_Excel_XLSX_file(Excel_file);
					if(Opened_file==XML) Save_XML (Excel_file);
					break;

				case IDM_SAVE_AS_XML:
					if(Opened_file==XLSX) Save_AsXML2(Excel_file);
					//	if(Opened_file==XLSX) Save_AsXML1(Excel_file);
					//if(Opened_file==XML) Save_LinkTable(Excel_file);
					break;

				case IDM_CLOSE:
					{
						Opened_file=0;
						memset(pT,0x00, MAX_ROW*sizeof(struct table));	
						for(int i=0;i<MAX_ROW;i++)
						{
							PutCell(hgrid2,i,1,"");PutCell(hgrid2,i,2,"");PutCell(hgrid2,i,3,"");PutCell(hgrid2,i,4,"");PutCell(hgrid2,i,5,"");PutCell(hgrid2,i,6,"");PutCell(hgrid2,i,7,"");PutCell(hgrid2,i,8,"");
						}
					}
					break;

				case  IDM_A: Table_Target_column=1;  	PutCell(hgrid2,1,Table_Target_column,Screenshot_col_name);InvalidateRect( hgrid2, NULL, TRUE );break;
				case  IDM_B: Table_Target_column=2;  	PutCell(hgrid2,1,Table_Target_column,Screenshot_col_name);InvalidateRect( hgrid2, NULL, TRUE );break;
				case  IDM_C: Table_Target_column=3;		PutCell(hgrid2,1,Table_Target_column,Screenshot_col_name);InvalidateRect( hgrid2, NULL, TRUE );break;
				case  IDM_D: Table_Target_column=4;  	PutCell(hgrid2,1,Table_Target_column,Screenshot_col_name);InvalidateRect( hgrid2, NULL, TRUE );break;
				case  IDM_E: Table_Target_column=5;		PutCell(hgrid2,1,Table_Target_column,Screenshot_col_name);InvalidateRect( hgrid2, NULL, TRUE );break;
				case  IDM_F: Table_Target_column=6;		PutCell(hgrid2,1,Table_Target_column,Screenshot_col_name);InvalidateRect( hgrid2, NULL, TRUE );break;
				case  IDM_G: Table_Target_column=7;		PutCell(hgrid2,1,Table_Target_column,Screenshot_col_name);InvalidateRect( hgrid2, NULL, TRUE );break;
				case  IDM_H: Table_Target_column=8;		PutCell(hgrid2,1,Table_Target_column,Screenshot_col_name);InvalidateRect( hgrid2, NULL, TRUE );break;


                case 500: //properties grid notification that something happened
                    {
                     if(HIWORD(wParam)==BGN_CELLCLICKED) //a cell was clicked in the properties grid
                         {
                          
                          int row,col,dtype;

                          //get the row and column of the clicked cell
                          row=LOWORD(lParam);
                          col=HIWORD(lParam);

                          //set the _BGCELL structure variable (cell) to this row and column
                          SetCell(&cell,row,col);
                          //get the data type that is in the cell
                          //in this instance, we're looking for BOOLEAN data (types 3 [TRUE] or 4 [FALSE])
                          //datatype 1 is alphanumeric data
                          //datatype 2 is numeric data
                          //datatype 3 is BOOLEAN TRUE data
                          //datatype 4 is BOOLEAN FALSE data
                          dtype=SendMessage(hgrid1,BGM_GETTYPE,(UINT)&cell,0);
                          if(dtype == 3) //bool true
                              {
                               //if the grid cell was true (checked checkbox), toggle it false
                               SendMessage(hgrid1,BGM_SETCELLDATA,(UINT)&cell,(long)"FALSE");
                               //send appropriate control message to the grid based
                               //on the row of the cell that was toggled

                               if(row==1)
                                   {
                                    SendMessage(hgrid2,BGM_SETALLOWCOLRESIZE,FALSE,0);
                                   }
                               if(row==2)
                                   {
                                    SendMessage(hgrid2,BGM_SETEDITABLE,FALSE,0);
                                   }
                               if(row==3)
                                   {
                                    SendMessage(hgrid2,BGM_SETELLIPSIS,FALSE,0);
                                   }
                               if(row==4)
                                   {
                                    SendMessage(hgrid2,BGM_SETCOLAUTOWIDTH,FALSE,0);
                                   }
                               if(row==5)
                                   {
                                    SendMessage(hgrid2,BGM_EXTENDLASTCOLUMN,FALSE,0);
                                   }
                               if(row==6)
                                   {
                                    SendMessage(hgrid2,BGM_SETCOLSNUMBERED,FALSE,0);
                                    LoadGrid2(hgrid2);
                                   }
                               if(row==7)
                                   {
                                    SendMessage(hgrid2,BGM_SETROWSNUMBERED,FALSE,0);
                                   }
                               if(row==8)
                                   {
                                    SendMessage(hgrid2,BGM_SHOWHILIGHT,FALSE,0);
                                   }
                               if(row==9)
                                   {
                                    SendMessage(hgrid2,BGM_SETCURSORCOLOR,(UINT)RGB(0,0,0),0);
                                   }
                               if(row==10)
                                   {
                                    SendMessage(hgrid2,BGM_SETGRIDLINECOLOR,(UINT)RGB(255,255,255),0);
                                   }
                              }
                          if(dtype == 4) //bool false
                              {
                               //if the grid cell was false (unchecked checkbox), toggle it true
                               SendMessage(hgrid1,BGM_SETCELLDATA,(UINT)&cell,(long)"TRUE");
                               //send appropriate control message to the grid based
                               //on the row of the cell that was toggled
                               if(row==1)
                                   {
                                    SendMessage(hgrid2,BGM_SETALLOWCOLRESIZE,TRUE,0);
                                   }
                               if(row==2)
                                   {
                                    SendMessage(hgrid2,BGM_SETEDITABLE,TRUE,0);
                                   }
                               if(row==3)
                                   {
                                    SendMessage(hgrid2,BGM_SETELLIPSIS,TRUE,0);
                                   }
                               if(row==4)
                                   {
                                    SendMessage(hgrid2,BGM_SETCOLAUTOWIDTH,TRUE,0);
                                   }
                               if(row==5)
                                   {
                                    SendMessage(hgrid2,BGM_EXTENDLASTCOLUMN,TRUE,0);
                                   }
                               if(row==6)
                                   {
                                    SendMessage(hgrid2,BGM_SETCOLSNUMBERED,TRUE,0);
                                    SendMessage(hgrid2,BGM_SETHEADERROWHEIGHT,21,0);
                                   }
                               if(row==7)
                                   {
                                    SendMessage(hgrid2,BGM_SETROWSNUMBERED,TRUE,0);
                                   }
                               if(row==8)
                                   {
                                    SendMessage(hgrid2,BGM_SHOWHILIGHT,TRUE,0);
                                   }
                               if(row==9)
                                   {
                                    SendMessage(hgrid2,BGM_SETCURSORCOLOR,(UINT)RGB(255,255,255),0);
                                   }
                               if(row==10)
                                   {
                                    SendMessage(hgrid2,BGM_SETGRIDLINECOLOR,(UINT)RGB(220,220,220),0);
                                   }
                              }
                         }
                    }
                    break;

				default:
				   return DefWindowProc(hWnd, message, wParam, lParam);
			}
			break;
		case WM_PAINT:
			hdc = BeginPaint(hWnd, &ps);
			// TODO: Add any drawing code here...
			EndPaint(hWnd, &ps);
			break;

        case WM_SIZE:
            {
              RECT rect;
              GetClientRect(hWnd,&rect);
              MoveWindow(hScreenshotWnd,rect.right-rect.right/3,0,rect.right/3,rect.bottom,TRUE);
              MoveWindow(hgrid2,0,0,rect.right-rect.right/3,rect.bottom,TRUE);
            }
            break;

        case WM_CREATE:
             RegisterGridClass(hInst); //initializes BABYGRID control
              Opened_file=0;                         //only call this function once in your program 
			DragAcceptFiles(hWnd,1);
			 memset(UserName,0,sizeof(UserName));
			GetUserName((LPTSTR)UserName,&User_name_len);
			UserName[27]=0;
			GetModuleFileName(hInst,ProgramPatch,sizeof(ProgramPatch));
			 {
				 for(int d=MAX_PATH;d>0;d--)
					if(ProgramPatch[d]==0x5c)
						{
							ProgramPatch[d+1]=0;
							break;
						}
			 }
				
			SYSTEMTIME sm;
			memset(&sm,0,sizeof(SYSTEMTIME));
			GetLocalTime(&sm);
			if(sm.wYear>2017||sm.wMonth>11 &&sm.wDay>2)
				{	
					MessageBoxW(hWnd, L"The debug version has expired. Please update the tool.", L"Error", MB_ICONSTOP);
					break;
				}

			 hThread_Sender2 = CreateThread( NULL, 0, Sender2, NULL, 0, &dwId );

             //create 2 grids for placement on the application main window
             //the 2 grids are placed in the WM_SIZE handler.
            // hgrid1=CreateWindowEx(WS_EX_CLIENTEDGE,"BABYGRID","Grid Properties",
            //     WS_VISIBLE|WS_CHILD,0,0,0,0,hWnd,(HMENU)500,hInst,NULL);
			 RECT rect;
              GetClientRect(hWnd,&rect);
			  hScreenshotWnd=CreatePreviewWindow(hWnd, rect.right-rect.right/3,0,rect.right/3,rect.bottom);

             hgrid2=CreateWindowEx(WS_EX_CLIENTEDGE,"BABYGRID","ScreenshotViewer",
                 WS_VISIBLE|WS_CHILD,0,0,0,0,hWnd,(HMENU)501,hInst,NULL);

             //Set grid2 (the working demonstration grid) to be 100 rows by 5 columns
             SendMessage(hgrid2,BGM_SETGRIDDIM,MAX_ROW,MAX_COLUMN);

             //set grid1 (the properties grid) to automatically size columns 
             //based on the length of the text entered into the cells
             SendMessage(hgrid1,BGM_SETCOLAUTOWIDTH,TRUE,0);
             //only want 2 columns, rows will be added as data is entered programmatically
             SendMessage(hgrid1,BGM_SETGRIDDIM,0,2);
             //I don't want a row header, so make it 0 pixels wide
             SendMessage(hgrid1,BGM_SETCOLWIDTH,0,0);
             //this grid won't use column headings, set header row height = 0
             SendMessage(hgrid1,BGM_SETHEADERROWHEIGHT,0,0);

			// SendMessage(hgrid2,BGM_SETEDITABLE,TRUE,0);
             //populate grid1 with data
             //LoadGrid1(hgrid1);
             //populate grid2 with initial demo data
             LoadGrid2(hgrid2);
             //make grid2 header row to initial height of 21 pixels
             SendMessage(hgrid2,BGM_SETHEADERROWHEIGHT,21,0);

			//// SendMessage(hgrid2,BGM_SETCOLAUTOWIDTH,TRUE,0);
			 //SendMessage(hgrid2,BGM_SETCOLSNUMBERED,FALSE,0);
			 SendMessage(hgrid2,BGM_SETALLOWCOLRESIZE,TRUE,0);

			// SendMessage(hgrid2,BGM_SETEDITABLE,TRUE,0);
			 
			 //strcpy(Excel_file,ProgramPatch);
			 //strcpy(Pictures_path,ProgramPatch);

				 // strcat(Excel_file,"MCN-BU.xlsx");
				 //Open_Excel_XLSX_file(Excel_file);

			 

			pT=0;
			pT=(struct table*)new struct table[MAX_ROW];
			if(pT==0)
			{
				MessageBoxW(hWnd, L"Cannot allocate memory for pT. Please close a unused programs and try again.", L"Warning", MB_ICONWARNING);			
				PostQuitMessage(0);	
			}
			memset(pT,0x00, MAX_ROW*sizeof(struct table));	
			//item_count=0;


            break;


		case WM_DESTROY:
			delete(pT);
			PostQuitMessage(0);
			break;


		case WM_DROPFILES:
			{
				char SourceFile[MAX_PATH]="";
				HDROP hdrop;
				hdrop=(struct HDROP__ *)wParam;

			
			
				int uNumFiles = DragQueryFile ( hdrop, -1, NULL, 0 );

				  if(uNumFiles>1)
				  {
			 		MessageBox(hWnd, "Sorry, this function is performing in debug mode.\n Only one file is supported.", "Warning", MB_OK);
			  		DragFinish ( hdrop );
					break;
				  }

				 
				for ( int uFile = 0; uFile < uNumFiles; uFile++ )
					{
					// Get the next filename from the HDROP info.
					if ( DragQueryFile ( hdrop, uFile, SourceFile, MAX_PATH ) > 0 )
						{
						// ***
						// Do whatever you want with the filename in szNextFile.
						// ***
						}
					}
				// Free up memory.
				DragFinish ( hdrop );
				//SetFocus(hgrid2);
				
				int j, filenamelen=strlen(SourceFile);
				for( j=filenamelen;j>0;j--)
				{
					if(SourceFile[j]=='.')
						if(SourceFile[j+1]=='x')if(SourceFile[j+2]=='m')if(SourceFile[j+3]=='l')
						{
							SendMessage(hWnd,WM_COMMAND,IDM_CLOSE,0);
							char* pFileName=0;
							strcpy(Excel_file,SourceFile);
								 
							strcpy(Pictures_path,Excel_file);
							pFileName=strrchr(Pictures_path,'\\');
							pFileName++;
							*pFileName=0;					
								 
								 
							{
								Open_XMLfile2 (Excel_file); Opened_file=XML ;
								pFileName--;
								*pFileName=0;
								pFileName=strrchr(Pictures_path,'\\');
								if(pFileName!=0)
								{
								pFileName++;
								*pFileName=0;
								}
								Table_Target_column=3;
							   Load_LinkTable(Excel_file);
							}
							break;
						}

						if(SourceFile[j]=='.')
						if(SourceFile[j+1]=='x')if(SourceFile[j+2]=='l')if(SourceFile[j+3]=='s')if(SourceFile[j+4]=='x')
						{
							SendMessage(hWnd,WM_COMMAND,IDM_CLOSE,0);
							
							char* pFileName=0;
							strcpy(Excel_file,SourceFile);
								 
							strcpy(Pictures_path,Excel_file);
							pFileName=strrchr(Pictures_path,'\\');
							pFileName++;
							*pFileName=0;					
								 						 
							{
								Open_Excel_XLSX_file(Excel_file); Opened_file=XLSX ;

								char szSTR[80];
										for(int h=0;h<MAX_COLUMN;h++)
										{memset(szSTR,0,80);
											GetCell( hgrid2, 2,h, szSTR);
											if(!memcmp(szSTR,"file:",strlen("file:")))
												Table_Target_column=h;
										}

								strcpy(Pictures_path,Excel_file);
								pFileName=strrchr(Pictures_path,'\\');
								pFileName++;
								*pFileName=0;	

							 Word_Count();
								
							   //Load_LinkTable(Excel_file);
							}
							break;
						}
					if(SourceFile[j]=='\\')
					{
						if(Table_Target_column==-1)
						{
							MessageBox(hWnd, "Please select a column for screenshot first in Menu->Screenshot", "Warning", MB_OK);
							break;
						}
						
						int currow=GetGurRow();
						if(currow>=1)
						{
							PutCell(hgrid2,currow,Table_Target_column,&SourceFile[j+1]);
							strcpy(pT[currow].pic,&SourceFile[j+1]);

							DrawPicture();
							SetForegroundWindow(hWnd);
							SetFocus(hWnd);
							SetFocus(hgrid2);
						}
						else
							{
			 				MessageBox(hWnd, "Please select a row first.", "Warning", MB_OK);
			  				//DragFinish ( hdrop );
							//break;
							}

						break;
					}	
				}
			} 
				break;

		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
   }
   return 0;
}

// Mesage handler for about box.
LRESULT CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
HDC hdc;
PAINTSTRUCT ps;
char  szSTR[300];
HFONT hCurFont;
int y;
	switch (message)
	{
			case WM_PAINT:
			                      
			hdc = BeginPaint(hDlg, &ps);
		y=0;
			hCurFont=CreateFontW(16,0,0,0,FW_NORMAL,0,0,0,
				RUSSIAN_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,
				PROOF_QUALITY,VARIABLE_PITCH|FF_ROMAN,L"Arial Unicode MS");
				
			SelectObject(hdc,hCurFont);

			SetBkMode(hdc, 0);


			
			if((ver%100)<10)
				wsprintf(szSTR,"Screenshot Viewer ver. %d.0%d %s beta",ver/100,ver%100,subver);
			else
				wsprintf(szSTR,"Screenshot Viewer ver. %d.%d %s beta",ver/100,ver%100,subver);
			TextOut(hdc, 90,20+y*20,szSTR,strlen(szSTR));
				y++;
 

				wsprintf(szSTR,"Copyright (c) 2002-2015 by");
			TextOut(hdc, 90,20+y*20,szSTR,strlen(szSTR)); 
				y++;

					wsprintf(szSTR,"David Hillard (mudcat@mis.net)" );
			TextOut(hdc, 90,20+y*20,szSTR,strlen(szSTR));
				y++;

					wsprintf(szSTR,"Sasha_p  (sasha_p@asus.com)" );
			TextOut(hdc, 90,20+y*20,szSTR,strlen(szSTR));
				y++;
 
 

		
			

			DeleteObject(hCurFont);

			EndPaint(hDlg, &ps);

			break;
			
			case WM_INITDIALOG:
				return TRUE;

		case WM_COMMAND:
			if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) 
			{
				EndDialog(hDlg, LOWORD(wParam));
				return TRUE;
			}
			break;
	}
    return FALSE;
}







void PutCell(HWND hgrid,int row, int col, char* text)
    {
     //worker function to keep from having to send hundreds of SendMessage() with
     //BGM_SETCELLDATA in the main program.  Just simplifies the main program code
     SetCell(&cell,row,col);
     SendMessage(hgrid,BGM_SETCELLDATA,(UINT)&cell,(long)text);
    }

void GetCell(HWND hgrid,int row, int col, char* text)
    {
     SetCell(&cell,row,col);
	SendMessage(hgrid,BGM_GETCELLDATA,(UINT)&cell,(long)text);
}

void LoadGrid2(HWND hgrid)
    {
     //load grid 2 with initial demo data
        //PutCell(hgrid,0,1,"Multi-line\nHeadings\nSupported");
        //PutCell(hgrid,0,2,"\n\nName");
        //PutCell(hgrid,0,3,"\n\nAge");

        SendMessage(hgrid,BGM_SETPROTECT,TRUE,0);
        //every cell entered after a BGM_SETPROTECT TRUE will set the 
        //protected attribute of that cell.  This keeps an editable grid
        //from allowing the user to overwrite whatever is in the protected cell

        SendMessage(hgrid,BGM_SETPROTECTCOLOR,(UINT)RGB(210,210,210),0);
        //the setprotectcolor is optional, but it gives a visual indication
        //of which cells are protected.

        //now put some data in the cells in grid2
        //PutCell(hgrid,1,2,"David");
        //PutCell(hgrid,2,2,"Maggie");
        //PutCell(hgrid,3,2,"Chester");
        //PutCell(hgrid,4,2,"Molly");
        //PutCell(hgrid,5,2,"Bailey");
                             
        //PutCell(hgrid,1,3,"43");
        //PutCell(hgrid,2,3,"41");
        //PutCell(hgrid,3,3,"3");
        //PutCell(hgrid,4,3,"3");
        //PutCell(hgrid,5,3,"1");

		//char buf[32]=""; GetCell(hgrid,1,2,buf);

        //PutCell(hgrid,10,5,"Shaded cells are write-protected.");

        SendMessage(hgrid,BGM_SETPROTECT,FALSE,0);
        //turn off automatic cell protection
        //if you don't turn off automatic cell protection, if the 
        //grid is editable, the user can enter data into empty cells
        //but cannot change what he entered... not good.

        //PutCell(hgrid,1,0,"Row Headers customizable");
 
    }

void LoadGrid1(HWND hgrid)
    {
     //load data into the properties grid

     PutCell(hgrid,1,1,"User Column Resizing");
     PutCell(hgrid,1,2,"FALSE");
     PutCell(hgrid,2,1,"User Editable");
     PutCell(hgrid,2,2,"FALSE");
     PutCell(hgrid,3,1,"Show Ellipsis");
     PutCell(hgrid,3,2,"TRUE");
     PutCell(hgrid,4,1,"Auto Column Size");
     PutCell(hgrid,4,2,"FALSE");
     PutCell(hgrid,5,1,"Extend Last Column");
     PutCell(hgrid,5,2,"TRUE");
     PutCell(hgrid,6,1,"Numbered Columns");
     PutCell(hgrid,6,2,"TRUE");
     PutCell(hgrid,7,1,"Numbered Rows");
     PutCell(hgrid,7,2,"TRUE");
     PutCell(hgrid,8,1,"Highlight Row");
     PutCell(hgrid,8,2,"TRUE");
     PutCell(hgrid,9,1,"Show Cursor");
     PutCell(hgrid,9,2,"TRUE");
     PutCell(hgrid,10,1,"Show Gridlines");
     PutCell(hgrid,10,2,"TRUE");

     //make the grid notify the program that the row in the 
     //grid has changed.  Usually this is done by the user clicking
     //a cell, or moving thru the grid with the keyboard.  But we
     //want the grid to initially send this message to get things going.
     //If we didn't call BGM_NOTIFYROWCHANGED, the first row would be 
     //hilighted, but the ACTION wouldn't be performed.

     SendMessage(hgrid,BGM_NOTIFYROWCHANGED,0,0);

     //make the properties grid have the focus when the application starts
     SetFocus(hgrid);

    }


