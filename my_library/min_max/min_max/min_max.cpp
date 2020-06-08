#include <iostream>


using std::cout;
using std::endl;
using std::cin;


int main()
{
	const short ARR_SIZE = 5;
	int arr[ARR_SIZE] = {2, 5, 6, 1, 3 };
	int min = arr[0], max = arr[0];
	for (int i=1; i<ARR_SIZE; ++i) 
	{
		if (min > arr[i]) min = arr[i];
		if (max < arr[i]) max = arr[i];
	}
	cout << "min = " << min << endl;
	cout << "max = " << max << endl;
	cin.get();
	return 0;
}