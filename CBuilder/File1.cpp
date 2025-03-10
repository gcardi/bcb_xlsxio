#include <iostream>
#include <tchar.h>
#include <stdlib.h>
#include <stdio.h>

// See original repositories for further info:
//   https://github.com/brechtsanders/xlsxio
//   https://github.com/madler/zlib
//   https://github.com/staticlibs/external_zlib
//   https://getitnow.embarcadero.com/expat/

#include <conio.h>

//#define PROCESS_FROM_MEMORY
//#define PROCESS_FROM_FILEHANDLE

#include <string.h>
#include <windows.h>
#include <io.h>
#include <sys/stat.h>
#include <fcntl.h>
#ifdef PROCESS_FROM_FILEHANDLE
# include <sys/types.h>
#endif

// The following items must be on the IDE's include search path
#include "xlsxio_write.h"
#include "xlsxio_read.h"

#if !defined(XML_UNICODE_WCHAR_T) && !defined(XML_UNICODE)
//UTF-8 version
#define X(s) s
#define XML_Char_printf printf
#else
//UTF-16 version
#define X(s) L##s
#define XML_Char_printf wprintf
#endif

// These items must be on the IDE's library search path
#pragma comment( lib, "xlsxio" )

const char* filename = "example.xlsx";

int write()
{
    xlsxiowriter handle;

    // open .xlsx file for writing (will overwrite if it already exists)
    if ( (handle = xlsxiowrite_open( filename, "MySheet" ) ) == nullptr ) {
        fprintf( stderr, "Error creating .xlsx file\n" );
        return 1;
    }

    // set row height
    xlsxiowrite_set_row_height(handle, 1);

    // how many rows to buffer to detect column widths
    xlsxiowrite_set_detection_rows(handle, 10);

    // write column names
    xlsxiowrite_add_column( handle, "Col1",  0 );
    xlsxiowrite_add_column( handle, "Col2", 21 );
    xlsxiowrite_add_column( handle, "Col3",  0 );
    xlsxiowrite_add_column( handle, "Col4",  2 );
    xlsxiowrite_add_column( handle, "Col5",  0 );
    xlsxiowrite_add_column( handle, "Col6",  0 );
    xlsxiowrite_add_column( handle, "Col7",  0 );
    xlsxiowrite_next_row(handle);

    // write data
    for ( int i = {}; i < 1000 ; ++i ) {
        xlsxiowrite_add_cell_string( handle, "Test" );
        xlsxiowrite_add_cell_string( handle, "A b  c   d    e     f\nnew line" );
        xlsxiowrite_add_cell_string( handle, "&% <test> \"'" );
        xlsxiowrite_add_cell_string( handle, nullptr );
        xlsxiowrite_add_cell_int( handle, i);
        xlsxiowrite_add_cell_datetime( handle, time( nullptr ) );
        xlsxiowrite_add_cell_float( handle, M_PI );
        xlsxiowrite_next_row( handle );
    }

    //close .xlsx file
    xlsxiowrite_close(handle);
    return 0;
}

int read()
{
    xlsxioreader xlsxioread;

    XML_Char_printf( X( "XLSX I/O library version %s\n" ), xlsxioread_get_version_string() );

#if defined( PROCESS_FROM_MEMORY )

    int filehandle;
    char* buf = nullptr;
    size_t buflen = 0;
    if ( ( filehandle = open( filename, O_RDONLY | O_BINARY ) ) != -1 ) {
        struct stat fileinfo;
        if ( fstat( filehandle, &fileinfo ) == 0 ) {
            if ( ( buf = malloc(fileinfo.st_size ) ) != nullptr ) {
                if ( fileinfo.st_size > 0 && read( filehandle, buf, fileinfo.st_size ) == fileinfo.st_size ) {
                    buflen = fileinfo.st_size;
                }
            }
        }
        close( filehandle );
    }

    if ( !buf || buflen == 0 ) {
        fprintf( stderr, "Error loading .xlsx file\n" );
        return 1;
    }

    if ( ( xlsxioread = xlsxioread_open_memory (buf, buflen, 1 ) ) == nullptr ) {
        fprintf( stderr, "Error processing .xlsx data\n" );
        return 1;
    }

#elif defined( PROCESS_FROM_FILEHANDLE )

    //open .xlsx file for reading
    int filehandle;

    if ( ( filehandle = _open( filename, O_RDONLY | O_BINARY, 0 ) ) == -1 ) {
        fprintf( stderr, "Error opening .xlsx file\n" );
        return 1;
    }

    if ( ( xlsxioread = xlsxioread_open_filehandle( filehandle ) ) == nullptr ) {
        fprintf( stderr, "Error reading .xlsx file\n" );
        return 1;
    }
#else
    // open .xlsx file for reading
    if ( ( xlsxioread = xlsxioread_open( filename ) ) == nullptr ) {
        fprintf( stderr, "Error opening .xlsx file\n" );
        return 1;
    }
#endif

    //list available sheets
    xlsxioreadersheetlist sheetlist;
    const XLSXIOCHAR* sheetname;
    printf( "Available sheets:\n" );
    if ( ( sheetlist = xlsxioread_sheetlist_open( xlsxioread ) ) != nullptr ) {
        while ( ( sheetname = xlsxioread_sheetlist_next( sheetlist ) ) != nullptr ) {
            XML_Char_printf( X(" - %s\n" ), sheetname );
        }
        xlsxioread_sheetlist_close( sheetlist );
    }

    //read values from first sheet
    XLSXIOCHAR* value;
    printf( "Contents of first sheet:\n" );
    xlsxioreadersheet sheet =
        xlsxioread_sheet_open( xlsxioread, nullptr, XLSXIOREAD_SKIP_EMPTY_ROWS );
    while ( xlsxioread_sheet_next_row( sheet ) ) {
        while ( ( value = xlsxioread_sheet_next_cell( sheet ) ) != NULL ) {
            XML_Char_printf( X( "%s\t" ), value );
            xlsxioread_free( value );
        }
        printf( "\n" );
    }
    xlsxioread_sheet_close( sheet );

    //clean up
    xlsxioread_close( xlsxioread );
    return 0;
}

int _tmain( int argc, _TCHAR* argv[] )
{
#ifdef _WIN32
    // switch Windows console to UTF-8
    SetConsoleOutputCP( CP_UTF8 );
#endif

    auto ret = write();
    if ( ret ) {
        return ret;
    }
    ret = read();

    printf( "Press any key\n" );
    getch();
    return ret;
}


