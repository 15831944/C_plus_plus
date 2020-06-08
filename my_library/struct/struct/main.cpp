#include <iostream>

struct Point {
int x;
int y;
};
struct {
Point top_left;
Point bottom_right;
}rect;

int main() {
rect.top_left.x = 0;
rect.top_left.y = 0;
rect.bottom_right.x = 100;
rect.bottom_right.y = 100;
std::cout << rect.top_left.x << " "
		<< rect.top_left.y << std::endl
		<< rect.bottom_right.x << " "
		<< rect.bottom_right.x << std::endl;
std::cin.get();
return 0;
}