#include <string.h>
#include "ExcelFile.h"

ExcelFile::ExcelFile()
{
	f = NULL;
}

ExcelFile::ExcelFile(char *fname)
{
	open(fname);
}

ExcelFile::~ExcelFile()
{
	if(f) close();
}

bool ExcelFile::open(char *fname)
{
	unsigned short xlheader[] = { 0x0809, 0x08, 0x00, 0x10, 0x00, 0x00 };
	
	f = fopen(fname, "wb");
	if(!f) return false;

	fwrite(xlheader, sizeof(unsigned short), 6, f);
	return true;
}

void ExcelFile::close()
{
	unsigned short xlfooter[] = { 0x0a, 0x00 };

	fwrite(xlfooter, sizeof(unsigned short), 2, f);
	fclose(f);
	f = NULL;
}

void ExcelFile::writeCell(unsigned short col, unsigned short row, int value)
{
	unsigned short xlcell[] = { 0x027e, 0x0a, 0x00, 0x00, 0x00 };

	xlcell[2] = row;
	xlcell[3] = col;

	fwrite(xlcell, 2, 5, f);

	int v = (value << 2) | 2;
	fwrite(&v, 4, 1, f);
}

void ExcelFile::writeCell(unsigned short col, unsigned short row, double value)
{
	unsigned short xlcell[] = { 0x203, 0x0e, 0x00, 0x00, 0x00 };

	xlcell[2] = row;
	xlcell[3] = col;

	fwrite(xlcell, 2, 5, f);
	fwrite(&value, 8, 1, f);
}

void ExcelFile::writeCell(unsigned short col, unsigned short row, char *value)
{
	unsigned short xlcell[] = { 0x0204, 0, 0, 0, 0, 0 };
	unsigned short len = strlen(value);

	xlcell[1] = 8 + len;
	xlcell[2] = row;
	xlcell[3] = col;
	xlcell[5] = len;

	fwrite(xlcell, 2, 6, f);
	fwrite(value, 1, len, f);
}
