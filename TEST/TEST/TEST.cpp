#include <iostream>
#include <locale.h>
using namespace std;
int main()
{
setlocale(0,"");
cout << "Выберите действия:"
<<"1: Сложить два числа"
<<"2: Умножить два числа"
<<"3: Найти разность двух чисел"
<< "4: Разделить два числа" <<endl;
int n;
cin>>n;//вводим идентификатор действий
//для нашего удобства введем сюда иницализацию двух чисел
double val1,val2;

switch(n)
{
case 1:
	cout<< "Введите число 1" <<endl;
	cin>>val1;
	cout<< "Введите число 2" <<endl;
	cin>>val2;
	cout<< "Сумму равна = "<<val1+val2<<endl;
break;

case 2:
	cout<< "Введите число 1" <<endl;
	cin>>val1;
	cout<< "Введите число 2" <<endl;
	cin>>val2;
	cout<< "Произведение равна = "<<val1*val2<<endl;
break;

case 3:
	cout<< "Введите число 1" <<endl;
	cin>>val1;
	cout<< "Введите число 2" <<endl;
	cin>>val2;
	cout<< "Разность равна = "<<val1-val2<<endl;
break;

case 4:
	cout<< "Введите число 1" <<endl;
	cin>>val1;
	cout<< "Введите число 2" <<endl;
	cin>>val2;
	cout<< "Деление равна = "<<val1/val2<<endl;
break;
//Остальные действия, я опущу их
default:
cout<< "Недопустимая операция" <<endl;
}

}