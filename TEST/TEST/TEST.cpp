#include <iostream>
#include <locale.h>
using namespace std;
int main()
{
setlocale(0,"");
cout << "�������� ��������:"
<<"1: ������� ��� �����"
<<"2: �������� ��� �����"
<<"3: ����� �������� ���� �����"
<< "4: ��������� ��� �����" <<endl;
int n;
cin>>n;//������ ������������� ��������
//��� ������ �������� ������ ���� ������������ ���� �����
double val1,val2;

switch(n)
{
case 1:
	cout<< "������� ����� 1" <<endl;
	cin>>val1;
	cout<< "������� ����� 2" <<endl;
	cin>>val2;
	cout<< "����� ����� = "<<val1+val2<<endl;
break;

case 2:
	cout<< "������� ����� 1" <<endl;
	cin>>val1;
	cout<< "������� ����� 2" <<endl;
	cin>>val2;
	cout<< "������������ ����� = "<<val1*val2<<endl;
break;

case 3:
	cout<< "������� ����� 1" <<endl;
	cin>>val1;
	cout<< "������� ����� 2" <<endl;
	cin>>val2;
	cout<< "�������� ����� = "<<val1-val2<<endl;
break;

case 4:
	cout<< "������� ����� 1" <<endl;
	cin>>val1;
	cout<< "������� ����� 2" <<endl;
	cin>>val2;
	cout<< "������� ����� = "<<val1/val2<<endl;
break;
//��������� ��������, � ����� ��
default:
cout<< "������������ ��������" <<endl;
}

}