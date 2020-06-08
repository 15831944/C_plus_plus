#include <stdio.h>

class ExcelFile
{
public:
	ExcelFile();
	ExcelFile(char *fname);
	~ExcelFile();

	bool open(char *fname);
	void writeCell(unsigned short col, unsigned short row, int value);
	void writeCell(unsigned short col, unsigned short row, double value);
	void writeCell(unsigned short col, unsigned short row, char *value);
	void close();

protected:
	FILE *f;
};
