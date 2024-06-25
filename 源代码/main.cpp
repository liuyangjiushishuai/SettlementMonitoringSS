

/*单水准闭合路线，点名之间不能有空格，简单闭合差的分配*/

#include"SettlementMonitoring.h"
//int main0()
//{
//	//读取数据文件
//	string data_path = "D:/陆上测量程序/水准闭合差分配/测试/XFH-改.txt";//原始文件路径
//	string save_path = "D:/陆上测量程序/水准闭合差分配/测试/结果/高程.txt"; //最终结果文件路径
//	vector<cz>cz_information;
//	readFile(data_path, cz_information);
//	//分配闭合差
//	allocationClosureError(save_path, cz_information);
//	return 0;
//}
int main()
{
	Project P(2);
	P.LoadData("D:/ce/ceshi.xlsx");
	P.calculateVariable();
	GenerateResults G;
	G.GenerateResultsTable(P, "D:/ce/result.xlsx");
	return 0;
}