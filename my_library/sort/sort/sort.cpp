#include <iostream>
#include <cstdlib>

int mysort(const void *arg1, const void *arg2);

int main() {
	const short ARR_SIZE = 5;
	int arr[ARR_SIZE] = {10, 5, 6, 1, 3};
	std::qsort(arr, ARR_SIZE, sizeof (int), mysort);
	for (int i=0; i<ARR_SIZE; ++i) {
		std::cout << arr[i] <<std::endl;
	}
	std::cin.get();
	return 0;
}
int mysort(const void *arg1, const void *arg2) {
	return *(int *)arg1 - *(int *)arg2;
}

//функция для сортировки по убыванию
//int mysort(const void *arg1, const void *arg2) {
//	return *(int *)arg1 - *(int *)arg2;
//}