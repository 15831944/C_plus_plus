#include <iostream>
#include <clocale>
#include <ctime>
#include <windows.h>
#include <conio.h>

int main() {

	std::setlocale(LC_ALL, "Russian_Russia.1251");
	std::tm ptm;
	const short SIZE = 100;
	char str[SIZE] = {0};
	while(1) {
		std::time_t t = std::time(0);
		int err = localtime_s(&ptm, &t); //получаем текущее время

		if (err) {
			std::cout << "Error" << std::endl;
			std::exit(1);
			}
			err = std::strftime(str, SIZE,
				"Сегодня: \n %A %d %b %Y %H:%M:%S \n %d.%m.%Y", &ptm);
		if (!err) {
			std::cout << "Error" << std::endl;
			std::exit(1);
			}
		std::cout << str << std::endl;
		std::cout << "   для выхода нажмите \"ESC\"" << std::endl;
		if(_kbhit() && getch() == 27) {
			break;
		}
		Sleep (1000); // "Засыпаем" на 1 секунду
		system("cls");
		};
	return 0;
}