#include <stdio.h>
#include "ExcelFile.h"

int main(int argc, char *argv[])
{
	ExcelFile xls;

	//printf("Creating .xls...\n");

	//if(!xls.open("C:\\test.xls"))
	if(!xls.open("C:\\1.xls"))
	{
		printf("Error");
		return 1;
	}

	//printf("writting data...\n");

	//xls.writeCell(5, 5, 123);      
	//xls.writeCell(1, 0, 123);
	/*xls.writeCell(0, 1, "Double value:");    xls.writeCell(1, 1, 3.14);
	xls.writeCell(0, 2, "Text value:");      xls.writeCell(1, 2, "Some text...");*/

	//printf("closing file...\n");

	xls.close();

	//printf("Done.\n");
	return 0;
}
