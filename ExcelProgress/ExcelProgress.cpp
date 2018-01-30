#include <iostream>
#include <string>
#include <vector>
#include <windows.h>
#include "libxl.h"
#include <algorithm>


#define COL_BARCODE   2
#define COL_NAME      3
#define COL_ID        4
#define COL_UNIT      8
#define COL_POSITION  16
#define COL_NUMBER    11

using namespace libxl;
using namespace std;

typedef struct Goods
{
public:
	string barcode;
	long id;
	string name;
	string unit;
	string position;
	int num;

}Goods;

bool SortByName(Goods m, Goods n)
{
	if (m.name < n.name)
		return true;
	else
		return false;
}

bool SortByPos(Goods m, Goods n)
{
	if (m.position < n.position)
		return true;
	else
		return false;
}

int main()
{
	cout << "开始处理 勿关!!!" << endl;

	Book* book = xlCreateXMLBook();
	vector<Goods> goodsList;
	if (book)
	{
		const char * x = "Halil Kural";
		const char * y = "windows-2723210a07c4e90162b26966a8jcdboe";
		book->setKey(x, y);
		if (book->load("../1.xlsx"))
		{
			Sheet* sheet = book->getSheet(0);
			if (sheet)
			{
				int row = 0;
				int col = 0;
				int rowNums = sheet->lastRow();
				int colNums = sheet->lastCol();

				//读取
				for (int i = 1; i<rowNums; ++i)
				{
					Goods temp;
					if(sheet->cellType(i, COL_BARCODE)!= CELLTYPE_EMPTY)
						temp.barcode  = sheet->readStr(i, COL_BARCODE);
					temp.id       = sheet->readNum(i, COL_ID);
					temp.name     = sheet->readStr(i, COL_NAME);
					temp.position = sheet->readStr(i, COL_POSITION);
					temp.unit     = sheet->readStr(i, COL_UNIT);
					temp.num      = sheet->readNum(i, COL_NUMBER);
					goodsList.push_back(temp);
				}

				//处理
				sort(goodsList.begin(), goodsList.end(), SortByName);
				for (vector<Goods>::iterator iter = goodsList.begin() + 1; iter != goodsList.end(); ++iter)
				{
					if (iter->name == (iter - 1)->name)
					{
						(iter - 1)->num += iter->num;
						iter = goodsList.erase(iter) - 1;
					}
				}
				sort(goodsList.begin(), goodsList.end(), SortByPos);
				sheet->clear(0, rowNums,0, colNums);

				//保存
				sheet->writeStr(0, 0, "商品条码");
				sheet->writeStr(0, 1, "商品名称");
				sheet->writeStr(0, 2, "商品单位");
				sheet->writeStr(0, 3, "商品ID");
				sheet->writeStr(0, 4, "库位");
				sheet->writeStr(0, 5, "库存数量");
				for (int i = 1; i < goodsList.size()+1; ++i)
				{
					sheet->writeStr(i, 0, (goodsList[i - 1].barcode).c_str());
					sheet->writeStr(i, 1, (goodsList[i - 1].name).c_str());
					sheet->writeStr(i, 2, (goodsList[i - 1].unit).c_str());
					sheet->writeNum(i, 3, (goodsList[i - 1].id));
					sheet->writeStr(i, 4, (goodsList[i - 1].position).c_str());
					sheet->writeNum(i, 5, (goodsList[i - 1].num));
				}
			}

			if (book->save("../1.xlsx"))
			{
				cout << "处理完毕" << endl; cout << "处理完毕" << endl;
				::ShellExecute(NULL, "open", "../1.xlsx", NULL, NULL, SW_SHOW);
			}
			else
			{
				cout << book->errorMessage() << endl;
			}
		}
		else
		{
			cout << "找不到文件 !" << endl;
			while (1);
		}

		book->release();
	}
	return 0;
}