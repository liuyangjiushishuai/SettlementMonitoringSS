#ifndef SETTLEMENTMONITORING_H_
#define SETTLEMENTMONITORING_H_





#include<iostream>
#include<string>
#include<vector>
#include <fstream>
#include <iomanip>
#include<cctype>
#include <xlnt/xlnt.hpp>
#include <wchar.h>
#include <Windows.h>
#include<unordered_map>
#include<unordered_set>
#include<algorithm>
#include<stdlib.h>
#include<locale.h>
#include <codecvt>
#include<Stringapiset.h>



using namespace std;
using namespace xlnt;
const string KnownPointName = "c02";    //已知点名称
const double KnownNumberHeight = 33.83300;//已知点高程
const double SettlementRateLimit = 0.04;  //沉降速率限差
const vector<string>RegionName = { "连云港新富海仓储有限公司旗台作业区液体化工品罐区工程桩基础构筑物沉降观测成果表",
"连云港新富海仓储有限公司旗台作业区液体化工品罐区工程非桩基础构筑物沉降观测成果表" };    //区域名称
const string ProjectName = "连云港新富海仓储有限公司旗台作业区液体化工品罐区工程沉降观测";  //工程名称
const string CompanyName = "连云港港口工程设计研究院有限公司";                          //公司名称
const string NewestMeasureData = "";                //最新一期的测量日期
const vector<string>NewAddPoint = {};   //新加的点
const vector<string>NewAddPointBelongGZW = {};  //新加的点所属的构筑物
const vector<int>NewAddPointBelongQY = {};    //新加的点所属的区域
const vector<string>TableOrder = {" ","一","二","三","四","五","六","七","八","九",
"十","十一","十二","十三","十四","十五","十六","十七","十八","十九","二十"};        //表序

//将日期化为年、月、日
vector<int>splitData(const string data);

//计算时间间隔
int calculateTimeInterval(const vector<int>f, const vector<int>b);

//UTF-8编码格式字符串 转普通sting类型
std::string UTF8_To_string(const std::string& str);


//Unicode 转 utf-8
bool WStringToString(const std::wstring& wstr, std::string& str);

//普通字符串转宽字符串
std::wstring to_wstring(const std::string& str);

//普通字符串转UTF-8字符串
std::string String_To_UTF8(const std::string& str);

//保留三位小数
double SaveThreeDecimal(double OriginalData,string&str);

//保留二位小数
double SaveTwoDecimal(double OriginalData,string&str);
//测站
class cz
{
public:
	string b_point;
	string f_point;
	double first_b_num;
	double first_f_num;
	double second_f_num;
	double second_b_num;
	double first_b_distance;
	double first_f_distance;
	double second_f_distance;
	double second_b_distance;
	double front_point_height;
	double difference_v;
	cz()
	{

	}
	cz(const string b_point, const string f_point, const double first_b_num,
		const double first_f_num, const double second_f_num, const double second_b_num,
		const double first_b_distance, const double first_f_distance, const double second_f_distance,
		const double second_b_distance)
	{
		this->b_point = b_point;
		this->f_point = f_point;
		this->first_b_num = first_b_num;
		this->first_f_num = first_f_num;
		this->second_f_num = second_f_num;
		this->second_b_num = second_b_num;
		this->first_b_distance = first_b_distance;
		this->first_f_distance = first_f_distance;
		this->second_f_distance = second_f_distance;
		this->second_b_distance = second_b_distance;
	}
	double calculateUnDistributionDifference()
	{
		double res = ((first_b_num - first_f_num) + (second_b_num - second_f_num)) / 2;
		return res;
	}
	double calculate_b_distance()
	{
		return (first_b_distance + second_b_distance) / 2;
	}
	double calculate_f_distance()
	{
		return (first_f_distance + second_f_distance) / 2;
	}
	double calculate_v(const double difference_km)
	{
		double res = (((calculate_b_distance() + calculate_f_distance())) / 1000) * difference_km;
		this->difference_v = res;
		return res;
	}
	double calculae_front_point_height(const double BackPointHeight)
	{
		double res = BackPointHeight + calculateUnDistributionDifference() + difference_v;
		this->front_point_height = res;
		return res;
	}
};
//沉降点
class cj_point
{
public:
	vector<string> cl_data; //测量日期
	vector<double> cl_height; //测量高程
	int Cl_Frequency;    //测量次数
	vector<double> SettlementAmount;  //沉降量
	vector<double> AccumulateSettlementAmount;  //累积沉降量
	vector<double> SettlementSpeed;  //沉降速度
	string name;     //沉降点的名称

	cj_point(int n, string name)
	{
		this->Cl_Frequency = n;
		this->name = name;
	}
	//计算沉降量
	vector<double>calculateSettlementAmount()
	{
		vector<double>res;
		for (int i = 0; i < Cl_Frequency; i++)
		{
			if (i == 0)
			{
				if (cl_height[i] != -100)
				{
					res.push_back(0);
				}
				else
				{
					res.push_back(-1000);
				}
			}
			else
			{
				if (cl_height[i] == -100)
				{
					res.push_back(-1000);
				}
				else
				{
					bool check = false;
					for (int j = i - 1; j >= 0; j--)
					{
						if (cl_height[j] == -100)
						{
							continue;
						}
						else
						{
							check = true;
							res.push_back(cl_height[i] - cl_height[j]);
							break;
						}
					}
					if (!check)
					{
						res.push_back(0);
					}
				}
			}
		}
		this->SettlementAmount = res;
		return res;
	}
	//计算累积沉降量
	vector<double>calculateASettlementAmount()
	{
		vector<double>res;
		for (int i = 0; i < Cl_Frequency; i++)
		{
			if (i == 0)
			{
				if (cl_height[i] != -100)
				{
					res.push_back(0);
				}
				else
				{
					res.push_back(-1000);
				}
			}
			else
			{
				if (cl_height[i] == -100)
				{
					res.push_back(-1000);
				}
				else
				{
					double CurrentASettlementAmount = 0;
					CurrentASettlementAmount += this->SettlementAmount[i];
					for (int j = i - 1; j >= 0; j--)
					{
						if (cl_height[j] != -100)
						{
							CurrentASettlementAmount += res[j];
							break;
						}
					}
					res.push_back(CurrentASettlementAmount);
				}
			}
		}
		this->AccumulateSettlementAmount = res;
		return res;
	}

	//计算沉降速度
	vector<double>calculateSettlementSpeed()
	{
		vector<double>res;
		for (int i = 0; i < Cl_Frequency; i++)
		{
			if (i == 0)
			{
				if (cl_height[i] != -100)
				{
					res.push_back(0);
				}
				else
				{
					res.push_back(-1000);
				}
			}
			else
			{
				if (cl_height[i] == -100)
				{
					res.push_back(-1000);
				}
				else
				{
					string f_time = "";
					string b_time = cl_data[i];
					for (int j = i - 1; j >= 0; j--)
					{
						if (cl_height[j] != -100)
						{
							f_time = cl_data[j];
							break;
						}
					}
					if (f_time == "")
					{
						res.push_back(0);
					}
					else
					{
						auto f_time_result = splitData(f_time);
						auto b_time_result = splitData(b_time);
						//计算时间间隔
						double interval_time = calculateTimeInterval(f_time_result, b_time_result);
						//计算沉降速度
						res.push_back(SettlementAmount[i] / interval_time);
					}
				}
			}
		}
		this->SettlementSpeed = res;
		return res;
	}
};
//构筑物
class GZW
{
public:
	string name;    //构筑物名称
	int SettlementPointNum; //包含的沉降点数
	int UnstablePointNum;     //未稳定的点数
	int frequency;        //该构筑物测量期数
	double LatestMaxAccumulateSettlementAmount;      //最后一期最大累积沉降量
	double LatestMinAccumulateSettlementAmount;      //最后一期最小累积沉降量
	double LatestMaxSettlementRate;                //最后一期最大沉降速度
	double LatestMinSettlementRate;               //最后一期最小沉降速度
	vector<cj_point>ContainSettlementPoint; //包含的沉降点
	vector<double>AverageSettlementAmount; //每期的平均沉降量
	vector<double>AverageSettlementRate;   //每期的平均沉降速度
	vector<double>AverageAccumulateSettlementAmount; //每期的平均累积沉降量


	GZW(string n/*构筑物名称*/, int f/*测量期数*/)
	{
		this->name = n;
		this->frequency = f;
	}
	//计算每期的平均沉降量
	void calculateEachIssueAverageSettlementAmount()
	{
		for (int i = 0; i < frequency; i++)
		{
			double CurrentAllSettlementAmount = 0;
			int CurrentSettlementPointNum = 0;
			for (int j = 0; j < ContainSettlementPoint.size(); j++)
			{
				if (ContainSettlementPoint[j].SettlementAmount[i] != -1000)
				{
					CurrentSettlementPointNum++;
					CurrentAllSettlementAmount += ContainSettlementPoint[j].SettlementAmount[i];
				}
			}
			if (CurrentSettlementPointNum == 0)
			{
				AverageSettlementAmount.push_back(-1000);
			}
			else
			{
				AverageSettlementAmount.push_back(CurrentAllSettlementAmount / CurrentSettlementPointNum);
			}
		}
	}
	//计算每期的平均沉降速度
	void calculateEachIssueAverageSettlementSpeed()
	{
		for (int i = 0; i < frequency; i++)
		{
			double CurrentSettlementSpeed = 0;
			int CurrentSettlePointNum = 0;
			for (int j = 0; j < ContainSettlementPoint.size(); j++)
			{
				if (ContainSettlementPoint[j].SettlementSpeed[i] != -1000)
				{
					CurrentSettlePointNum++;
					CurrentSettlementSpeed += ContainSettlementPoint[j].SettlementSpeed[i];
				}
			}
			if (CurrentSettlePointNum == 0)
			{
				AverageSettlementRate.push_back(-1000);
			}
			AverageSettlementRate.push_back(CurrentSettlementSpeed / CurrentSettlePointNum);
		}
	}
	//计算每期的平均累积沉降量
	void calcualteEachIssueAverageAccumulateSettlementAmount()
	{
		for (int i = 0; i < frequency; i++)
		{
			double CurrentAccumulateSettlementAmount = 0;
			int CurrentSettlementPointNum = 0;
			for (int j = 0; j < ContainSettlementPoint.size(); j++)
			{
				if (ContainSettlementPoint[j].AccumulateSettlementAmount[i] != -1000)
				{
					CurrentSettlementPointNum++;
					CurrentAccumulateSettlementAmount += ContainSettlementPoint[j].AccumulateSettlementAmount[i];
				}
			}
			if (CurrentSettlementPointNum == 0)
			{
				AverageAccumulateSettlementAmount.push_back(-1000);
			}
			else
			{
				AverageAccumulateSettlementAmount.push_back(CurrentAccumulateSettlementAmount / CurrentSettlementPointNum);
			}
		}
	}
	//计算最后一期最大累积沉降量
	double calcualteMaxAccumulateSettlementAmount()
	{
		double CompareValue = 0;
		double res = 0;
		int count = 0;
		for (int i = 0; i < ContainSettlementPoint.size(); i++)
		{
			double CurrentAccumulateSettlementAmount = ContainSettlementPoint[i].AccumulateSettlementAmount[frequency - 1];
			if (CurrentAccumulateSettlementAmount == -1000)
			{
				count++;
				continue;
			}
			if (fabs(fabs(CurrentAccumulateSettlementAmount) - CompareValue) > 1e-6 && fabs(CurrentAccumulateSettlementAmount) > CompareValue)
			{
				CompareValue = fabs(CurrentAccumulateSettlementAmount);
				res = CurrentAccumulateSettlementAmount;
			}
		}
		if (count == ContainSettlementPoint.size())
		{
			this->LatestMaxAccumulateSettlementAmount = -1000;
			return -1000;
		}
		else
		{
			this->LatestMaxAccumulateSettlementAmount = res;      //最后一期最大累积沉降量
		}
		return res;
	}
	//计算最后一期的最小累积沉降量
	double calcualteMinAccumulateSettlementAmount()
	{
		double CompareValue = 1000000;
		double res = 1000000;
		int count = 0;
		for (int i = 0; i < ContainSettlementPoint.size(); i++)
		{
			double CurrentAccumulateSettlementAmount = ContainSettlementPoint[i].AccumulateSettlementAmount[frequency - 1];
			if (CurrentAccumulateSettlementAmount == -1000)
			{
				count++;
				continue;
			}
			if (fabs(fabs(CurrentAccumulateSettlementAmount) - CompareValue) > 1e-6 && fabs(CurrentAccumulateSettlementAmount) < CompareValue)
			{
				CompareValue = fabs(CurrentAccumulateSettlementAmount);
				res = CurrentAccumulateSettlementAmount;
			}
		}
		if (count == ContainSettlementPoint.size())
		{
			this->LatestMinAccumulateSettlementAmount = -1000;
			return -1000;
		}
		else
		{
			this->LatestMinAccumulateSettlementAmount = res;      //最后一期最小累积沉降量
		}
		return res;
	}
	//计算最后一期的最大沉降速度
	double calcualteMaxSettlementRate()
	{
		double CompareValue = 0;
		double res = 0;
		int count = 0;
		for (int i = 0; i < ContainSettlementPoint.size(); i++)
		{
			double CurrentSettlementRate = ContainSettlementPoint[i].SettlementSpeed[frequency - 1];
			if (CurrentSettlementRate == -1000)
			{
				count++;
				continue;
			}
			if (fabs(fabs(CurrentSettlementRate) - CompareValue) > 1e-6 && fabs(CurrentSettlementRate) > CompareValue)
			{
				CompareValue = fabs(CurrentSettlementRate);
				res = CurrentSettlementRate;
			}
		}
		if (count == ContainSettlementPoint.size())
		{
			this->LatestMaxSettlementRate = -1000;                //最后一期最大沉降速度
		}
		else
		{
			this->LatestMaxSettlementRate = res;                //最后一期最大沉降速度
		}
		return res;
	}
	//计算最后一期的最小沉降速度
	double calcualteMinSettlementRate()
	{
		double CompareValue = 1000000;
		double res = 1000000;
		int count = 0;
		for (int i = 0; i < ContainSettlementPoint.size(); i++)
		{
			double CurrentSettlementRate = ContainSettlementPoint[i].SettlementSpeed[frequency - 1];
			if (CurrentSettlementRate == -1000)
			{
				count++;
				continue;
			}
			if (fabs(fabs(CurrentSettlementRate) - CompareValue) > 1e-6 && fabs(CurrentSettlementRate) < CompareValue)
			{
				CompareValue = fabs(CurrentSettlementRate);
				res = CurrentSettlementRate;
			}
		}
		if (count == ContainSettlementPoint.size())
		{
			this->LatestMinSettlementRate = -1000;
		}
		else
		{
			this->LatestMinSettlementRate = res;
		}
		return res;
	}
	//统计未达到稳定标准的点数
	int countUnstablePointNum()
	{
		int res = 0;
		if (this->SettlementPointNum!=0)
		{
			int count = 0;
			for (int i = 0; i < this->ContainSettlementPoint.size(); i++)
			{
				double CurrentSettlementRate = this->ContainSettlementPoint[i].SettlementSpeed[frequency - 1];
				if (abs(CurrentSettlementRate + 1000) <= 1e-6)
				{
					count++;
					continue;
				}
				if (abs(abs(CurrentSettlementRate) - SettlementRateLimit) > 1e-6 && abs(CurrentSettlementRate) >
					SettlementRateLimit)
				{
					res++;
				}
			}
			if (count == this->SettlementPointNum)
			{
				res = 0;
			}
		}
		this->UnstablePointNum = res;
		return res;
	}
};
//区域
class QY
{
public:
	string name;  //区域名称
	string EngineeringName;  //工程名称
	string CompanyName;       //公司名称
	int frequency;      //该区域所测期数
	vector<GZW>ContainGZW;//包含的构筑物
	vector<string>data;   //该区域所测的日期
	double MaxAverageAccumulateSettlementAmount; //最大平均累计沉降量
	double MinAverageAccumulateSettlementAmount; //最小平均累计沉降量
	string MaxAASettlementAmountGZW;     //最大平均累计沉降量构筑物
	string MinAASettlementAmountGZW;    //最小平均累计沉降量构筑物
	double MaxAverageSettlementRate; //最大平均沉降速度
	double MinAverageSettlementRate; //最小平均沉降速度
	string MaxASettlementRateGZW;     //最大平均沉降速度构筑物
	string MinASettlementRateGZW;    //最小平均沉降速度构筑物
	string MaxSettlementRatePoint;   //最大沉降速度点
	double MaxSettlementRate;     //最大沉降速度
	string MaxSettlementRateGZW;   //最大沉降速度构筑物
	int MoreLimitGZWNum;        //超限的构筑物数量

	QY(string n/*区域名称*/, string EN/*工程名称*/, string CN/*公司名称*/,int frequency/*该区域所测期数*/)
	{
		this->name = n;
		this->EngineeringName = EN;
		this->CompanyName = CN;
		this->frequency = frequency;
	}

	//计算最大的平均累积沉降量
	void calculateMaxAverageAccumulateSettlementAmount()
	{
		double CompareValue = 0;
		double res = 0;
		int count = 0;
		string s;  //最大平均累计沉降量所对应的构筑物
		for (int i = 0; i < ContainGZW.size(); i++)
		{
			double rate = ContainGZW[i].frequency;   //该构筑物所测频率
			double CurrentAverageAccumulateSettlementAmount = ContainGZW[i].AverageAccumulateSettlementAmount[rate - 1];
			if (abs(CurrentAverageAccumulateSettlementAmount + 1000) < 1e-6)
			{
				count++;
				continue;
			}
			if (fabs(fabs(CurrentAverageAccumulateSettlementAmount) - CompareValue) >= 1e-6 &&
				abs(CurrentAverageAccumulateSettlementAmount) > CompareValue)
			{
				res = CurrentAverageAccumulateSettlementAmount;
				CompareValue = abs(CurrentAverageAccumulateSettlementAmount);
				s = ContainGZW[i].name;
			}
		}
		if (count == ContainGZW.size())
		{
			this->MaxAverageAccumulateSettlementAmount = -1000;
			this->MaxAASettlementAmountGZW = "None";
		}
		else
		{
			this->MaxAverageAccumulateSettlementAmount = res;
			this->MaxAASettlementAmountGZW = s;
		}
	}
	//计算最小的平均累计沉降量
	void calculateMinAverageAccumulateSettlementAmount()
	{
		double res = 1000000;
		double CompareValue = 1000000;
		string s;  //最小平均累计沉降量所对应的构筑物
		int count = 0;
		for (int i = 0; i < ContainGZW.size(); i++)
		{
			double rate = ContainGZW[i].frequency;   //该构筑物所测频率
			double CurrentAverageAccumulateSettlementAmount = ContainGZW[i].AverageAccumulateSettlementAmount[rate - 1];
			if (abs(CurrentAverageAccumulateSettlementAmount + 1000) < 1e-6)
			{
				count++;
				continue;
			}
			if (fabs(fabs(CurrentAverageAccumulateSettlementAmount) - CompareValue) >= 1e-6 &&abs(CurrentAverageAccumulateSettlementAmount) <
				CompareValue)
			{
				res = CurrentAverageAccumulateSettlementAmount;
				CompareValue = abs(CurrentAverageAccumulateSettlementAmount);
				s = ContainGZW[i].name;
			}
		}
		if (count == ContainGZW.size())
		{
			s = "None";
			res = -1000;
		}
		this->MinAverageAccumulateSettlementAmount = res;
		this->MinAASettlementAmountGZW = s;
	}
	//计算最大平均沉降速度
	void calculateMaxAverageSettlementRate()
	{
		double res = 0;
		double CompareValue = 0;
		int count = 0;
		string s;  //最大平均沉降速度所对应的构筑物
		for (int i = 0; i < ContainGZW.size(); i++)
		{
			double rate = ContainGZW[i].frequency;   //该构筑物所测频率
			double CurrentAverageSettlementRate = ContainGZW[i].AverageSettlementRate[rate - 1];
			if (abs(CurrentAverageSettlementRate + 1000) < 1e-6)
			{
				count++;
				continue;
			}
			if (fabs(fabs(CurrentAverageSettlementRate) - CompareValue) >= 1e-6 &&
				abs(CurrentAverageSettlementRate) > CompareValue)
			{
				res = CurrentAverageSettlementRate;
				CompareValue = abs(CurrentAverageSettlementRate);
				s = ContainGZW[i].name;
			}
		}
		if (count == ContainGZW.size())
		{
			res = -1000;
			s = "None";
		}
		this->MaxAverageSettlementRate = res;
		this->MaxASettlementRateGZW = s;
	}
	//计算最小平均沉降速度
	void calculateMinAverageSettlementRate()
	{
		double res = 1000000;
		double CompareValue = 1000000;
		string s;  //最小平均沉降速度所对应的构筑物
		int count = 0;
		for (int i = 0; i < ContainGZW.size(); i++)
		{
			double rate = ContainGZW[i].frequency;   //该构筑物所测频率
			double CurrentAverageSettlementRate = ContainGZW[i].AverageSettlementRate[rate - 1];
			if (abs(CurrentAverageSettlementRate + 1000) < 1e-6)
			{
				count++;
				continue;
			}
			if (fabs(fabs(CurrentAverageSettlementRate) - CompareValue) >= 1e-6 && abs(CurrentAverageSettlementRate) <
				CompareValue)
			{
				res = CurrentAverageSettlementRate;
				CompareValue = abs(CurrentAverageSettlementRate);
				s = ContainGZW[i].name;
			}
		}
		if (count == ContainGZW.size())
		{
			res = -1000;
			s = "None";
		}
		this->MinAverageSettlementRate = res;
		this->MinASettlementRateGZW = s;
	}
	//计算最大沉降速度
	void calculateMaxSettlementRate()
	{
		double res = 0;
		double CompareValue = 0;
		string s;  //最大沉降速度点
		string s2; //最大沉降速度构筑物
		int count = 0;
		for (int i = 0; i < ContainGZW.size(); i++)
		{
			int ContainPointNum = ContainGZW[i].ContainSettlementPoint.size(); //该构筑物所包含的点数
			int frequency = ContainGZW[i].frequency;    //该构筑物所测的期数
			if (ContainPointNum == 0)
			{
				continue;
			}
			for (int j = 0; j < ContainPointNum; j++)
			{
				double CurrentSettlementRate = ContainGZW[i].ContainSettlementPoint[j].SettlementSpeed[frequency - 1];
				if (abs(CurrentSettlementRate + 1000) < 1e-6)
				{
					continue;
				}
				if (abs(abs(CurrentSettlementRate)- CompareValue) >1e-6&& abs(CurrentSettlementRate)> CompareValue)
				{
					res = CurrentSettlementRate;
					CompareValue = abs(CurrentSettlementRate);
					s = ContainGZW[i].ContainSettlementPoint[j].name;
					s2 = ContainGZW[i].name;
				}
				count++;
			}
		}
		if (count == 0)
		{
			s = "None";
			s2 = "None";
			res = -1000;
		}
		this->MaxSettlementRate = res;
		this->MaxSettlementRatePoint = s;
		this->MaxSettlementRateGZW = s2;
	}
	//统计超限的构筑物
	int countMoreLimitGZW()
	{
		int res = 0;
		for (int i = 0; i < ContainGZW.size(); i++)
		{
			double CurrentMaxSettlementRate = ContainGZW[i].calcualteMaxSettlementRate();
			if (abs(CurrentMaxSettlementRate+1000)<1e-6)
			{
				continue;
			}
			if (abs(abs(CurrentMaxSettlementRate)-SettlementRateLimit)>1e-6&&abs(CurrentMaxSettlementRate)> SettlementRateLimit)
			{
				res++;
			}
		}
		this->MoreLimitGZWNum = res;
		return res;
	}
};
//项目
class Project
{
public:
	vector<QY>qy; //所包含的区域
	int RegionNum;  //区域数量
	vector<unordered_set<string>>ContainGZW;  //所包含的构筑物
	vector<unordered_map<string, unordered_set<string>>>ContainCJP;//所包含的沉降点
	vector<unordered_map<string, vector<double>>>ContainHeight;  //沉降点所包含的高程
	vector<vector<string>>GZW_order;      //按顺序装的构筑物
	vector<unordered_map<string, vector<string>>>CJP_order;  //按顺序装的沉降点



	Project(int RN)
	{
		this->RegionNum = RN;
		/*this->ContainGZW.resize(this->RegionNum);
		this->ContainCJP.resize(this->RegionNum);
		this->ContainHeight.resize(this->RegionNum);*/
	}
	//加载沉降监测数据
	void LoadData(string path/*沉降监测数据文件路径*/)
	{
		xlnt::workbook wb;
		wb.load(path);


		//区域循环
		for (int i = 0; i < RegionNum; i++)
		{
			this->ContainGZW.push_back(unordered_set<string>());
			this->ContainCJP.push_back(unordered_map<string, unordered_set<string>>());
			this->ContainHeight.push_back(unordered_map<string, vector<double>>());
			this->GZW_order.push_back(vector<string>());
			this->CJP_order.push_back(unordered_map<string, vector<string>>());
			vector<GZW>structures;  //构筑物
			vector<string>time;
			auto ws = wb.sheet_by_index(i);
			int c1 = 0;
			int	StartQs = 0;   //记录每个表格的开始期数
			int SheetNum = 0;   //记录读到的表格数
			int CurrentSheetTitleRow = 0;   //记录目前的表标题行数
			string CurrentGZWName = "";     //目前的构筑物名称
			string CurrentCJPName = "";     //目前的沉降点名称
			for (auto row : ws.rows(false))
			{
				int c2 = 0;
				int c3 = StartQs;  //统计高程观测次数
				for (auto cell : row)
				{
					//读第一行，将该区域所有观测日期读入
					if (c1 == 0 && cell.has_value())
					{
						time.push_back(UTF8_To_string(cell.to_string()));
					}
					//读到观测日期这一行
					else if (UTF8_To_string(cell.to_string()) == "观测日期")
					{
						SheetNum++;
						if (SheetNum == 2)
						{
							StartQs += 3;
						}
						else if (SheetNum > 2)
						{
							StartQs += 2;
						}
						break;
					}
					//读到构筑物名称这一行
					else if (UTF8_To_string(cell.to_string()) == "构筑物名称")
					{
						CurrentSheetTitleRow = c1;
						break;
					}
					//存储构筑物
					else if (cell.has_value() && UTF8_To_string(ws.cell(c2 + 1, CurrentSheetTitleRow + 1).to_string()) == "构筑物名称")
					{
						CurrentGZWName = UTF8_To_string(cell.to_string());    //目前的构筑物名称
						if (this->ContainGZW[i].count(CurrentGZWName) == 0)
						{
							this->ContainGZW[i].emplace(CurrentGZWName);
							this->ContainCJP[i].emplace(CurrentGZWName, unordered_set<string>());
							this->GZW_order[i].push_back(CurrentGZWName);
						}
					}
					//存储沉降点
					else if (cell.has_value() && UTF8_To_string(ws.cell(c2 + 1, CurrentSheetTitleRow + 1).to_string()) == "点号")
					{
						CurrentCJPName = UTF8_To_string(cell.to_string()); //目前的沉降点名称
						if (this->ContainCJP[i][CurrentGZWName].count(CurrentCJPName) == 0)
						{
							this->ContainCJP[i][CurrentGZWName].emplace(CurrentCJPName);
							this->ContainHeight[i].emplace(CurrentCJPName, vector<double>());
							this->ContainHeight[i][CurrentCJPName].resize(time.size(), -100);
							this->CJP_order[i][CurrentGZWName].push_back(CurrentCJPName);
						}
					}
					//将沉降点高程加入
					else if (UTF8_To_string(ws.cell(c2 + 1, CurrentSheetTitleRow + 1).to_string()).find("值") != string::npos)
					{
						if (cell.has_value())
						{
							try
							{
								this->ContainHeight[i][CurrentCJPName][c3] = stod(UTF8_To_string(cell.to_string()));
							}
							catch (const std::exception&)
							{

							}
						}
						c3++;
					}
					c2++;
				}
				c1++;
			}
			//创建该区域的构筑物
			for (int j = 0; j < this->GZW_order[i].size(); j++)
			{
				string CurrentGZWName = this->GZW_order[i][j];
				GZW CurrentGZW(CurrentGZWName, time.size());
				CurrentGZW.SettlementPointNum = this->ContainCJP[i][CurrentGZWName].size();
				for (int k = 0; k < this->CJP_order[i][CurrentGZWName].size(); k++)
				{
					string CurrentCJPName = this->CJP_order[i][CurrentGZWName][k];
					cj_point CurrentCJP(time.size(), CurrentCJPName);
					CurrentCJP.cl_height = this->ContainHeight[i][CurrentCJPName];
					CurrentCJP.cl_data = time;
					CurrentGZW.ContainSettlementPoint.push_back(CurrentCJP);
				}
				structures.push_back(CurrentGZW);
			}
			//创建一个区域
			QY CurrentQY(RegionName[i], ProjectName, CompanyName,time.size());
			CurrentQY.ContainGZW = structures;
			CurrentQY.data = time;
			this->qy.push_back(CurrentQY);
		}
		//检测
		for (int i = 0; i < this->qy.size(); i++)
		{
			for (int j = 0; j < this->qy[i].ContainGZW.size(); j++)
			{
				cout << this->qy[i].ContainGZW[j].name << endl;
				for (int k = 0; k < this->qy[i].ContainGZW[j].ContainSettlementPoint.size(); k++)
				{
					cout << this->qy[i].ContainGZW[j].ContainSettlementPoint[k].name << " ";
					for (int m = 0; m < this->qy[i].ContainGZW[j].ContainSettlementPoint[k].cl_height.size(); m++)
					{
						cout << std::fixed << std::setprecision(3) << this->qy[i].ContainGZW[j].ContainSettlementPoint[k].cl_height[m] << " ";
					}
					cout << endl;
				}
			}
		}
	}
	//以规定的间隔符分割字符串
	void splitString(vector<string>& s, const char FSplitStr, const char SSplitStr, const string& str)
	{
		string temper;
		for (int i = 0; i < str.size(); i++)
		{
			if (str[i] == FSplitStr || str[i] == SSplitStr)
			{
				if (temper != "")
				{
					s.push_back(temper);
					temper.clear();
				}
			}
			else
			{
				temper += str[i];
			}
		}
		s.push_back(temper);
	}
	//读取测站信息
	void read_station_information(cz& station, const vector<string>station_line)
	{
		int flag = 0;
		for (int i = 0; i < station_line.size(); i++)
		{
			flag++;
			//分割字符串
			vector<string>s;
			splitString(s, '|', ' ', station_line[i]);
			//第一次后视
			if (flag == 1)
			{
				station.b_point = s[5];
				station.first_b_num = stod(s[9]);
				station.first_b_distance = stod(s[12]);
			}
			//第一次前视
			else if (flag == 2)
			{
				station.f_point = s[5];
				station.first_f_num = stod(s[9]);
				station.first_f_distance = stod(s[12]);
			}
			//第二次前视
			else if (flag == 3)
			{
				station.second_f_num = stod(s[9]);
				station.second_f_distance = stod(s[12]);
			}
			//第二次后视
			else if (flag == 4)
			{
				station.second_b_num = stod(s[9]);
				station.second_b_distance = stod(s[12]);
			}
		}
	}
	//读取水准原始数据
	void readFile(string path, vector<cz>& d)
	{
		//打开文件
		std::ifstream file(path);
		std::string line;
		vector<string>data;  //记录所有数据
		int flag = 0;//判断是不是一个测站的
		int line_num = 0;//记录数据行数
		vector<string>station_information; //记录一个测站的数据
		if (file.is_open())
		{
			while (std::getline(file, line))
			{
				line_num++;
				data.push_back(line);
				if (line_num <= 4)
				{
					continue;
				}
				if (line.find("#") != string::npos || line.find("Measurement repeated") != string::npos
					|| line.find("Station repeated") != string::npos || line.find("End-Line") != string::npos
					|| line.find("Cont-Line") != string::npos || line.find("Sh") != string::npos ||
					line.find("Db") != string::npos || line.find("Start-Line") != string::npos || line.find(":") == string::npos)
				{
					continue;
				}
				flag++;
				if (flag >= 1 && flag <= 4)
				{
					station_information.push_back(line);
				}
				else if (flag == 5)
				{
					//测试
					for (int i = 0; i < station_information.size(); i++)
					{
						cout << station_information[i] << endl;
					}
					//开始记入一个测站的信息
					cz station;
					read_station_information(station, station_information);
					d.push_back(station);  //将测站装入
					flag = 0;    //将测站计数置为0
					station_information.clear(); //将记录测站数据的容器置空
				}
				std::cout << line << std::endl;
			}
			////读取闭合差
			//vector<string>SplitedInformation;
			//splitString(SplitedInformation, '|', ' ', data[line_num - 3]);
			//difference = stod(SplitedInformation[6]);
			file.close();
		}
		else
		{
			std::cout << "无法打开要读取的文件" << std::endl;
		}
	}
	//计算闭合差
	double calculateClosureError(vector<cz>& d)
	{
		double res = 0;
		for (int i = 0; i < d.size(); i++)
		{
			res += d[i].calculateUnDistributionDifference();
		}
		return res;
	}
	//按路线长度分配闭合差
	void allocationClosureError(const string path, vector<cz>& d)
	{
		//计算水准路线总长度
		double all_length = 0;
		double AFDistance = 0;
		double ABDistance = 0;
		//计算闭合差
		double ClosureError = calculateClosureError(d);
		cout << "测站总数:" << d.size() << endl;
		//计算累积后视距离
		for (int i = 0; i < d.size(); i++)
		{
			ABDistance += d[i].calculate_b_distance();
		}
		//计算累积前视距离
		for (int i = 0; i < d.size(); i++)
		{
			AFDistance += d[i].calculate_f_distance();
		}
		cout << "累积后视距离:" << ABDistance << endl;
		cout << "累积前视距离" << AFDistance << endl;
		for (int i = 0; i < d.size(); i++)
		{
			all_length += d[i].calculate_b_distance() + d[i].calculate_f_distance();
		}
		//计算每千米闭合差
		double ClosureError_km = -ClosureError / (all_length / 1000);
		//计算每个测站的高差改正数
		for (int i = 0; i < d.size(); i++)
		{
			d[i].calculate_v(ClosureError_km);
			//计算每个点改正后的高程
			if (i == 0)
			{
				d[i].calculae_front_point_height(KnownNumberHeight);
			}
			else
			{
				d[i].calculae_front_point_height(d[i - 1].front_point_height);
			}
		}


		//将每个点的高程写入文本
		std::ofstream outfile(path); // 创建ofstream对象，并打开文件

		if (outfile.is_open()) // 检查文件是否成功打开	
		{
			outfile << d[0].b_point << "," << KnownNumberHeight << std::endl; // 将基准点写入
			//将待测点写入
			for (int i = 0; i < d.size(); i++)
			{
				outfile << d[i].f_point << "," << std::fixed << std::setprecision(3) << d[i].front_point_height << std::endl;
			}
			outfile.close(); // 关闭文件
		}
		else
		{
			std::cout << "无法打开要写入的文件" << std::endl;
		}
		cout << "计算完成" << endl;
	}
	//将最新一期数据加入项目
	void addNewestData(vector<cz>& d, string path1/*原始水准数据路径*/, string path2/*写入水准数据文件路径*/, int flag/*判断有没有新点,1：有，0：没有*/)
	{
		readFile(path1, d);
		allocationClosureError(path2, d);
		//存储最新一期的沉降点
		unordered_map<string, double>NewestCJP;
		for (int i = 0; i < d.size(); i++)
		{
			string name = d[i].f_point;
			double height = d[i].front_point_height;
			if (name == KnownPointName)
			{
				continue;
			}
			NewestCJP[name] = height;
		}
		for (int j = 0; j < RegionNum; j++)
		{
			for (int k = 0; k < this->qy[j].ContainGZW.size(); k++)
			{
				for (int m = 0; m < this->qy[j].ContainGZW[k].ContainSettlementPoint.size(); m++)
				{
					this->qy[j].ContainGZW[k].ContainSettlementPoint[m].cl_data.push_back(NewestMeasureData);
					this->qy[j].ContainGZW[k].ContainSettlementPoint[m].Cl_Frequency++;
					this->qy[j].ContainGZW[k].ContainSettlementPoint[m].cl_height.push_back(-100);
				}
				this->qy[j].ContainGZW[k].frequency++;
			}
		}
		for (int i = 0; i < d.size(); i++)
		{
			auto CurrentCJPName = d[i].f_point;
			auto CurrentHeight = d[i].front_point_height;
			for (int j = 0; j < RegionNum; j++)
			{
				for (int k = 0; k < this->qy[j].ContainGZW.size(); k++)
				{
					for (int m = 0; m < this->qy[j].ContainGZW[k].ContainSettlementPoint.size(); m++)
					{
						if (this->qy[j].ContainGZW[k].ContainSettlementPoint[m].name == CurrentCJPName)
						{
							this->qy[j].ContainGZW[k].ContainSettlementPoint[m].cl_height[this->qy[j].ContainGZW[k].frequency - 1] = CurrentHeight;
						}
					}
				}
			}
		}
		//如果有新点
		if (flag)
		{
			if (NewAddPoint.size() == 0)
			{
				cout << "请加入新加的点" << endl;
				system("pause");
				exit(1);
			}
			for (int i = 0; i < NewAddPoint.size(); i++)
			{
				string CJPName = NewAddPoint[i];
				int BelongQY = NewAddPointBelongQY[i];
				string BelongGZW = NewAddPointBelongGZW[i];
				cj_point CJP(this->qy[0].ContainGZW[0].frequency, CJPName);
				CJP.cl_data = this->qy[0].ContainGZW[0].ContainSettlementPoint[0].cl_data;
				CJP.cl_height.resize(this->qy[0].ContainGZW[0].frequency, -100);
				CJP.cl_height[this->qy[0].ContainGZW[0].ContainSettlementPoint[0].cl_data.size() - 1] = NewestCJP[CJPName];
				if (this->ContainGZW[BelongQY].count(BelongGZW) == 0)
				{
					GZW gzw(BelongGZW, this->qy[0].ContainGZW[0].frequency);
					gzw.ContainSettlementPoint.push_back(CJP);
					gzw.SettlementPointNum++;
					this->ContainGZW[BelongQY].emplace(BelongGZW);
					this->ContainCJP[BelongQY].emplace(BelongGZW, unordered_set<string>{CJPName});
					this->ContainHeight[BelongQY].emplace(CJPName, vector<double>(this->qy[0].ContainGZW[0].ContainSettlementPoint[0].cl_data.size(), -100));
					this->ContainHeight[BelongQY][CJPName][this->qy[0].ContainGZW[0].ContainSettlementPoint[0].cl_data.size() - 1] = NewestCJP[CJPName];
					this->GZW_order[BelongQY].push_back(BelongGZW);
					this->CJP_order[BelongQY].emplace(BelongGZW, vector<string>{CJPName});
					this->qy[BelongQY].ContainGZW.push_back(gzw);
				}
				else
				{
					this->ContainCJP[BelongQY][BelongGZW].emplace(CJPName);
					this->ContainHeight[BelongQY].emplace(CJPName, vector<double>(this->qy[0].ContainGZW[0].ContainSettlementPoint[0].cl_data.size(), -100));
					this->ContainHeight[BelongQY][CJPName][this->qy[0].ContainGZW[0].ContainSettlementPoint[0].cl_data.size() - 1] = NewestCJP[CJPName];
					this->CJP_order[BelongQY][BelongGZW].push_back(CJPName);
					for (int j = 0; j < this->qy[BelongQY].ContainGZW.size(); j++)
					{
						if (this->qy[BelongQY].ContainGZW[j].name == BelongGZW)
						{
							this->qy[BelongQY].ContainGZW[j].ContainSettlementPoint.push_back(CJP);
							this->qy[BelongQY].ContainGZW[j].SettlementPointNum++;
							break;
						}
					}
				}
			}
		}
	}
	//计算各种变量
	void calculateVariable()
	{
		//计算沉降量、累积沉降量、沉降速度
		for (int i = 0; i < this->qy.size(); i++)
		{
			for (int j = 0; j < this->qy[i].ContainGZW.size(); j++)
			{
				for (int k = 0; k < this->qy[i].ContainGZW[j].ContainSettlementPoint.size(); k++)
				{
					//计算沉降量
					this->qy[i].ContainGZW[j].ContainSettlementPoint[k].calculateSettlementAmount();
					//计算累积沉降量
					this->qy[i].ContainGZW[j].ContainSettlementPoint[k].calculateASettlementAmount();
					//计算沉降速度
					this->qy[i].ContainGZW[j].ContainSettlementPoint[k].calculateSettlementSpeed();
				}
			}
		}
		/*计算每期的平均沉降量、平均沉降速度、平均累积沉降量，最后一期的最大、最小累积沉降量、
		最后一期的最大、最小沉降速度、统计不稳定的沉降点数*/
		for (int i = 0; i < this->qy.size(); i++)
		{
			for (int j = 0; j < this->qy[i].ContainGZW.size(); j++)
			{
				//计算每期的平均沉降量
				this->qy[i].ContainGZW[j].calculateEachIssueAverageSettlementAmount();
				//计算每期的平均累积沉降量
				this->qy[i].ContainGZW[j].calcualteEachIssueAverageAccumulateSettlementAmount();
				//计算每期的平均沉降速度
				this->qy[i].ContainGZW[j].calculateEachIssueAverageSettlementSpeed();
				//计算最后一期的最大累积沉降量
				this->qy[i].ContainGZW[j].calcualteMaxAccumulateSettlementAmount();
				//计算最后一期的最小累积沉降量
				this->qy[i].ContainGZW[j].calcualteMinAccumulateSettlementAmount();
				//计算最后一期的最大沉降速度
				this->qy[i].ContainGZW[j].calcualteMaxSettlementRate();
				//计算最后一期的最小沉降速度
				this->qy[i].ContainGZW[j].calcualteMinSettlementRate();
				//统计不稳定的沉降点数
				this->qy[i].ContainGZW[j].countUnstablePointNum();
			}
		}
		//计算最大平均累积沉降量、计算最小平均累积沉降量、计算最大平均沉降速度、计算最小平均沉降速度、计算最大沉降速度、统计超限的构筑物
		for (int i = 0; i < qy.size(); i++)
		{
			//计算最大平均累积沉降量
			qy[i].calculateMaxAverageAccumulateSettlementAmount();
			//计算最小平均累积沉降量
			qy[i].calculateMinAverageAccumulateSettlementAmount();
			//计算最大平均沉降速度
			qy[i].calculateMaxAverageSettlementRate();
			//计算最小平均沉降速度
			qy[i].calculateMinAverageSettlementRate();
			//计算最大沉降速度
			qy[i].calculateMaxSettlementRate();
			//统计超限的构筑物
			qy[i].countMoreLimitGZW();
		}
		//测试
		string s1;
		for (int i = 0; i < this->qy.size(); i++)
		{
			for (int j = 0; j < this->qy[i].ContainGZW.size(); j++)
			{
				cout << this->qy[i].ContainGZW[j].name << endl;
				for (int k = 0; k < this->qy[i].ContainGZW[j].ContainSettlementPoint.size(); k++)
				{
					cout << this->qy[i].ContainGZW[j].ContainSettlementPoint[k].name << " ";
					for (int m = 0; m < this->qy[i].ContainGZW[j].ContainSettlementPoint[k].cl_height.size(); m++)
					{
						cout << this->qy[i].ContainGZW[j].ContainSettlementPoint[k].cl_height[m] << " " <<
							this->qy[i].ContainGZW[j].ContainSettlementPoint[k].SettlementAmount[m] << " " <<
							this->qy[i].ContainGZW[j].ContainSettlementPoint[k].AccumulateSettlementAmount[m] << " " << fixed << setprecision(3) <<
							SaveThreeDecimal(this->qy[i].ContainGZW[j].ContainSettlementPoint[k].SettlementSpeed[m],s1) << " ";
					}
					cout << endl;
				}
			}
		}
		//测试
		/*for (int i = 0; i < this->qy.size(); i++)
		{
			for (int j = 0; j < this->qy[i].ContainGZW.size(); j++)
			{
				for (int k = 0; k < this->qy[i].ContainGZW[j].AverageSettlementAmount.size(); k++)
				{
					cout << fixed << setprecision(2) << SaveTwoDecimal(this->qy[i].ContainGZW[j].AverageSettlementAmount[k]) << " " <<
						SaveTwoDecimal(this->qy[i].ContainGZW[j].AverageAccumulateSettlementAmount[k]) << " " << fixed << setprecision(3) <<
						SaveThreeDecimal(this->qy[i].ContainGZW[j].AverageSettlementRate[k]) << "  ";
				}
				cout << endl;
			}
		}*/
		//测试
		/*for (int i = 0; i < this->qy.size(); i++)
		{
			for (int j = 0; j < this->qy[i].ContainGZW.size(); j++)
			{
				cout << fixed << setprecision(2) << SaveTwoDecimal(this->qy[i].ContainGZW[j].LatestMaxAccumulateSettlementAmount)
					<< " " << SaveTwoDecimal(this->qy[i].ContainGZW[j].LatestMinAccumulateSettlementAmount) << " " <<
					fixed << setprecision(3) << SaveThreeDecimal(this->qy[i].ContainGZW[j].LatestMaxSettlementRate) << " " <<
					SaveThreeDecimal(this->qy[i].ContainGZW[j].LatestMinSettlementRate) << endl;
			}
		}*/
		//测试
		/*for (int i = 0; i < this->qy.size(); i++)
		{
			cout << i + 1 << "    " << "最大平均累积沉降量:" << qy[i].MaxAASettlementAmountGZW<<","<<fixed << setprecision(2) <<
				SaveTwoDecimal(qy[i].MaxAverageAccumulateSettlementAmount)
				<< "最小平均累积沉降量:" <<qy[i].MinAASettlementAmountGZW<<"," << SaveTwoDecimal(qy[i].MinAverageAccumulateSettlementAmount)
				<< "最大平均沉降速度:" <<qy[i].MaxASettlementRateGZW<<","<< fixed << setprecision(3)
				<< SaveThreeDecimal(qy[i].MaxAverageSettlementRate) << "最小平均沉降速度:" <<qy[i].MinASettlementRateGZW<<","<<
				SaveThreeDecimal(qy[i].MinAverageSettlementRate) <<
				"最大沉降速度:" <<qy[i].MaxSettlementRateGZW<<","<<qy[i].MaxSettlementRatePoint<<","<<SaveThreeDecimal(qy[i].MaxSettlementRate) 
				<< "超限构筑物数量:" << qy[i].MoreLimitGZWNum << endl;
		}*/
		//测试
		for (int i = 0; i < this->qy.size(); i++)
		{
			for (int j = 0; j < this->qy[i].ContainGZW.size(); j++)
			{
				cout << this->qy[i].ContainGZW[j].name << "    " << "总点数:" << this->qy[i].ContainGZW[j].
					SettlementPointNum << "  " << "超限的点数:" << this->qy[i].ContainGZW[j].UnstablePointNum << endl;
			}
		}
	}
	//计算间隔月数
	pair<int, int> calculateIntervalMonths()
	{

	}
};

//生成成果
class GenerateResults
{
public:
	
	GenerateResults()
	{

	}
	//生成成果表
	void GenerateResultsTable(const Project&P/*项目*/,const string& path/*保存路径*/)
	{
		xlnt::workbook wb;
		auto sheet1 = wb.active_sheet();
		sheet1.title("sheet0");
		for (int i = 0; i < P.qy.size(); i++)
		{
			string CurrentRegionName = RegionName[i]; //目前的区域名称
			auto ws = wb.copy_sheet(sheet1);
			int frequency = P.qy[i].frequency;   //该区域所测的期数
			vector<string>data = P.qy[i].data;
			int c = 0;     //记录表数
			int StartFrequency = 0;  //记录每个表的开始期数
			int CurrentFrequency = 0;   //记录当前的期数
			int StartRowNum = 0;  //记录每个表的开始行数
			while (CurrentFrequency<frequency)
			{
				c++;
				string CurrentTableName= CurrentRegionName +"（" + TableOrder[c] + "）"; //目前的表名
				int limit = 0;
				if (c == 1)
				{
					limit = CurrentFrequency + 3;
					//将表名写入
					ws.cell(1, StartRowNum+1).value(String_To_UTF8(CurrentTableName));
					//将工程名称写入
					ws.cell(3, StartRowNum + 2).value(String_To_UTF8(ProjectName));
					ws.cell(1, StartRowNum + 2).value(String_To_UTF8("工程名称"));
					//将观测日期写入
					ws.cell(1, StartRowNum + 4).value(String_To_UTF8("观测日期"));
					while (CurrentFrequency<min(limit,frequency))
					{
						if (CurrentFrequency == 0)
						{
							//将第一期日期写入
							ws.cell(3, StartRowNum + 4).value(String_To_UTF8(data[CurrentFrequency]));
						}
						else if (CurrentFrequency == 1)
						{
							//将第二期的日期写入
							ws.cell(4, StartRowNum + 4).value(String_To_UTF8(data[CurrentFrequency]));
						}
						else if (CurrentFrequency == 2)
						{
							//将第三期的日期写入
							ws.cell(7, StartRowNum + 4).value(String_To_UTF8(data[CurrentFrequency]));
						}
						CurrentFrequency++;
					}
					//写入构筑物名称
					ws.cell(1, StartRowNum + 6).value(String_To_UTF8("构筑物名称"));
					//写入点号
					ws.cell(2, StartRowNum + 6).value(String_To_UTF8("点号"));
					//写入表头
					for (int j = StartFrequency; j < min(limit,frequency); j++)
					{
						string TableHeader = "第" + TableOrder[j + 1] + "次观测值";
						if (j == 0)
						{
							ws.cell(3, StartRowNum + 6).value(String_To_UTF8(TableHeader));
							ws.cell(3, StartRowNum + 7).value("(mm)");
						}
						else if (j == 1)
						{
							ws.cell(4, StartRowNum + 6).value(String_To_UTF8(TableHeader));
							ws.cell(4, StartRowNum + 7).value("(mm)");
							ws.cell(5, StartRowNum + 6).value(String_To_UTF8("本次沉降量"));
							ws.cell(5, StartRowNum + 7).value("(mm)");
							ws.cell(6, StartRowNum + 6).value(String_To_UTF8("沉降速率"));
							ws.cell(6, StartRowNum + 7).value("(mm/d)");
						}
						else if (j == 2)
						{
							ws.cell(7, StartRowNum + 6).value(String_To_UTF8(TableHeader));
							ws.cell(7, StartRowNum + 7).value("(mm)");
							ws.cell(8, StartRowNum + 6).value(String_To_UTF8("本次沉降量"));
							ws.cell(8, StartRowNum + 7).value("(mm)");
							ws.cell(9, StartRowNum + 6).value(String_To_UTF8("累计沉降量"));
							ws.cell(9, StartRowNum + 7).value("(mm)");
							ws.cell(10, StartRowNum + 6).value(String_To_UTF8("沉降速率"));
							ws.cell(10, StartRowNum + 7).value("(mm/d)");
						}
					}
					//写入数据
					int GZWStartRow = StartRowNum + 7;   //记录每个构筑物的开始行数
					for (int j = 0; j < P.qy[i].ContainGZW.size(); j++)
					{
						//将构筑物的名称写入
						ws.cell(1, GZWStartRow + 1).value(String_To_UTF8(P.qy[i].ContainGZW[j].name));

						for (int m = 0; m < P.qy[i].ContainGZW[j].ContainSettlementPoint.size(); m++)
						{
							//将点名写入
							ws.cell(2, GZWStartRow+1 + m).value(String_To_UTF8(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].name));
							//将数据写入
							int f = 0;
							for (int k = StartFrequency; k < min(limit,frequency); k++)
							{
								if (k == 0)
								{
									if (abs(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k]+100)<=1e-6)
									{
										continue;
									}
									//将高程数据写入
									string NumStr = "";
									SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k], NumStr);
									ws.cell(3 + f, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
								}
								else if (k == 1)
								{
									if (abs(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k] + 100) <= 1e-6)
									{
										continue;
									}
									//将高程数据写入
									string NumStr = "";
									SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k], NumStr);
									ws.cell(3 + f, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
									//将本次沉降量写入
									SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].SettlementAmount[k], NumStr);
									ws.cell(3 + f+1, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
									//将沉降速率写入
									SaveThreeDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].SettlementSpeed[k], NumStr);
									ws.cell(3 + f+2, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
								}
								else if (k == 2)
								{
									if (abs(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k] + 100) <= 1e-6)
									{
										continue;
									}
									//将高程数据写入
									string NumStr = "";
									SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k], NumStr);
									ws.cell(3 + f+2, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
									//将本次沉降量写入
									SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].SettlementAmount[k], NumStr);
									ws.cell(3 + f + 3, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
									//将累计沉降量写入
									SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].AccumulateSettlementAmount[k], NumStr);
									ws.cell(3 + f + 4, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
									//将沉降速率写入
									SaveThreeDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].SettlementSpeed[k], NumStr);
									ws.cell(3 + f + 5, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
								}
								f++;
							}
						}
						GZWStartRow += P.qy[i].ContainGZW[j].ContainSettlementPoint.size(); //刷新构筑物开始行数
					}
					//将观测间隔天数写入
					ws.cell(1, GZWStartRow + 1).value(String_To_UTF8("观测间隔天数(d)"));
					int f = 0;
					for (int k = StartFrequency; k < min(limit, frequency); k++)
					{
						if (k == 0)
						{
							ws.cell(3 + f, GZWStartRow + 1).value("0");
						}
						else if (k == 1)
						{
							auto f_time = splitData(data[k-1]);
							auto b_time= splitData(data[k]);
							int interval = calculateTimeInterval(f_time, b_time);
							ws.cell(3 + f, GZWStartRow + 1).value(interval);
						}
						else if (k == 2)
						{
							auto f_time = splitData(data[k - 1]);
							auto b_time = splitData(data[k]);
							int interval = calculateTimeInterval(f_time, b_time);
							ws.cell(3 + f+2, GZWStartRow + 1).value(interval);
						}
						f++;
					}
					//将说明写入
					ws.cell(1, GZWStartRow + 2).value(String_To_UTF8("说明"));
					ws.cell(3, GZWStartRow + 2).value(String_To_UTF8("1、“-”表示下沉，“+”表示上升。"));
					//将监测单位写入
					string MonitoringUnit="监测单位：";
					MonitoringUnit += CompanyName;
					ws.cell(1, GZWStartRow + 4).value(String_To_UTF8(MonitoringUnit));
					//更新每个表的开始行数
					StartRowNum = GZWStartRow + 5;
					//更新每个表的开始期数
					StartFrequency = CurrentFrequency;
				}
				else
				{
					limit = CurrentFrequency + 2;
					//将表名写入
					ws.cell(1, StartRowNum + 1).value(String_To_UTF8(CurrentTableName));
					//将工程名称写入
					ws.cell(3, StartRowNum + 2).value(String_To_UTF8(ProjectName));
					ws.cell(1, StartRowNum + 2).value(String_To_UTF8("工程名称"));
					//将观测日期写入
					ws.cell(1, StartRowNum + 4).value(String_To_UTF8("观测日期"));
					int f = 0;
					while (CurrentFrequency < min(limit, frequency))
					{
						ws.cell(3+4*f, StartRowNum + 4).value(String_To_UTF8(data[CurrentFrequency]));
						CurrentFrequency++;
						f++;
					}
					f = 0;
					//写入构筑物名称
					ws.cell(1, StartRowNum + 6).value(String_To_UTF8("构筑物名称"));
					//写入点号
					ws.cell(2, StartRowNum + 6).value(String_To_UTF8("点号"));
					//写入表头
					for (int j = StartFrequency; j < min(limit, frequency); j++)
					{
						string TableHeader = "第" + TableOrder[j + 1] + "次观测值";
						ws.cell(3+4*f, StartRowNum + 6).value(String_To_UTF8(TableHeader));
						ws.cell(3+4*f, StartRowNum + 7).value("(mm)");
						ws.cell(3 + 4 * f+1, StartRowNum + 6).value(String_To_UTF8("本次沉降量"));
						ws.cell(3 + 4 * f + 1, StartRowNum + 7).value("(mm)");
						ws.cell(3+4*f+2, StartRowNum + 6).value(String_To_UTF8("累计沉降量"));
						ws.cell(3 + 4 * f + 2, StartRowNum + 7).value("(mm)");
						ws.cell(3+4*f+3, StartRowNum + 6).value(String_To_UTF8("沉降速率"));
						ws.cell(3+4*f+3, StartRowNum + 7).value("(mm/d)");
						f++;
					}
					//写入数据
					int GZWStartRow = StartRowNum + 7;   //记录每个构筑物的开始行数
					for (int j = 0; j < P.qy[i].ContainGZW.size(); j++)
					{
						//将构筑物的名称写入
						ws.cell(1, GZWStartRow + 1).value(String_To_UTF8(P.qy[i].ContainGZW[j].name));

						for (int m = 0; m < P.qy[i].ContainGZW[j].ContainSettlementPoint.size(); m++)
						{
							//将点名写入
							ws.cell(2, GZWStartRow + 1 + m).value(String_To_UTF8(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].name));
							//将数据写入
							int f = 0;
							for (int k = StartFrequency; k < min(limit, frequency); k++)
							{
								if (abs(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k] + 100) <= 1e-6)
								{
									continue;
								}
								//将高程数据写入
								string NumStr = "";
								SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k], NumStr);
								ws.cell(3 + 4*f, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
								//将本次沉降量写入
								SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].SettlementAmount[k], NumStr);
								ws.cell(3 + 4*f + 1, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
								//将累计沉降量写入
								SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].AccumulateSettlementAmount[k], NumStr);
								ws.cell(3 + 4*f + 2, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
								//将沉降速率写入
								SaveThreeDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].SettlementSpeed[k], NumStr);
								ws.cell(3 + 4*f + 3, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
								f++;
							}
						}
						GZWStartRow += P.qy[i].ContainGZW[j].ContainSettlementPoint.size(); //刷新构筑物开始行数
					}
					//将观测间隔天数写入
					ws.cell(1, GZWStartRow + 1).value(String_To_UTF8("观测间隔天数(d)"));
					f = 0;
					for (int k = StartFrequency; k < min(limit, frequency); k++)
					{
						auto f_time = splitData(data[k - 1]);
						auto b_time = splitData(data[k]);
						int interval = calculateTimeInterval(f_time, b_time);
						ws.cell(3 + 4*f, GZWStartRow + 1).value(interval);
						f++;
					}
					//将说明写入
					ws.cell(1, GZWStartRow + 2).value(String_To_UTF8("说明"));
					ws.cell(3, GZWStartRow + 2).value(String_To_UTF8("1、“-”表示下沉，“+”表示上升。"));
					//将监测单位写入
					string MonitoringUnit = "监测单位：";
					MonitoringUnit += CompanyName;
					ws.cell(1, GZWStartRow + 4).value(String_To_UTF8(MonitoringUnit));
				}

			}
		}
		wb.remove_sheet(sheet1);
		wb.save(path);
	}
	//生成成果分析表
	void GenerateResultAnalysisTable(const Project& P/*项目*/, const string& path/*保存路径*/)
	{
		xlnt::workbook wb;
		auto sheet1 = wb.active_sheet();
		sheet1.title("sheet0");
		for (int i = 0; i < P.qy.size(); i++)
		{
			auto ws = wb.copy_sheet(sheet1);
			int CurrentRegionPeriod = P.qy[i].frequency;
			vector<string>data = P.qy[i].data;   //目前区域所测的时间
			int record = 0;
			int StartRow = 1;  //每个表的开始行数
			for (int j = 0; j < P.qy[i].ContainGZW.size(); j++)
			{
				int CurrentPeriod = 1;
				while (CurrentPeriod< CurrentRegionPeriod)
				{
					//写入构筑物名称
					string CurrentGZWName = P.qy[i].ContainGZW[j].name;
					ws.cell(1, StartRow - 1 + 1).value(String_To_UTF8(CurrentGZWName));
					if (CurrentPeriod == 1)
					{
						//写入表头
						ws.cell(1, StartRow + 1).value(String_To_UTF8("点号"));
						ws.cell(2, StartRow + 1).value(String_To_UTF8(data[CurrentPeriod]));
						if (CurrentPeriod + 1 < CurrentRegionPeriod)
						{
							ws.cell(4, StartRow + 1).value(String_To_UTF8(data[CurrentPeriod+1]));
						}
						if (CurrentPeriod + 2 < CurrentRegionPeriod)
						{
							ws.cell(7, StartRow + 1).value(String_To_UTF8(data[CurrentPeriod + 2]));
						}
						ws.cell(2, StartRow + 2).value(String_To_UTF8("本次沉降量(mm)"));
						ws.cell(3, StartRow + 2).value(String_To_UTF8("沉降速率(mm/d)"));
						
						ws.cell(4, StartRow + 2).value(String_To_UTF8("本次沉降量(mm)"));
						ws.cell(5, StartRow + 2).value(String_To_UTF8("累计沉降量(mm)"));
						ws.cell(6, StartRow + 2).value(String_To_UTF8("沉降速率(mm/d)"));
						
						ws.cell(7, StartRow + 2).value(String_To_UTF8("本次沉降量(mm)"));
						ws.cell(8, StartRow + 2).value(String_To_UTF8("累计沉降量(mm)"));
						ws.cell(9, StartRow + 2).value(String_To_UTF8("沉降速率(mm/d)"));
					}
					else
					{
						ws.cell(1, StartRow + 1).value(String_To_UTF8("点号"));
						int f = 0;
						for (int i = CurrentPeriod; i < CurrentPeriod + 3; i++)
						{
							if (i < CurrentRegionPeriod)
							{
								ws.cell(2 + f * 3, StartRow + 1).value(String_To_UTF8(data[i]));
							}
							ws.cell(2 + f * 3, StartRow + 2).value(String_To_UTF8("本次沉降量(mm)"));
							ws.cell(2 + f * 3 + 1, StartRow + 2).value(String_To_UTF8("累计沉降量(mm)"));
							ws.cell(2 + f * 3 + 1 + 1, StartRow + 2).value(String_To_UTF8("沉降速率(mm/d)"));
							f++;
						}

					}
					int limit = CurrentPeriod + 3;
					for (int m = 0;m< P.qy[i].ContainGZW[j].ContainSettlementPoint.size(); m++)
					{
						//写入点名
						string CurrentPointName = P.qy[i].ContainGZW[j].ContainSettlementPoint[m].name;
						ws.cell(1, StartRow + 2 + m + 1).value(String_To_UTF8(CurrentPointName));
						int f = 0;
						while (CurrentPeriod<min(limit, CurrentRegionPeriod))
						{
							double CurrentSettlementAmount = P.qy[i].ContainGZW[j].ContainSettlementPoint[m].
								SettlementAmount[CurrentPeriod];  //目前的沉降量
							double CurrentASettleAmount = P.qy[i].ContainGZW[j].ContainSettlementPoint[m].
								AccumulateSettlementAmount[CurrentPeriod];    //目前的累计沉降量
							double CurrentSettleSpeed = P.qy[i].ContainGZW[j].ContainSettlementPoint[m].
								SettlementSpeed[CurrentPeriod];  //目前的沉降速度
							if (CurrentPeriod == 1)
							{
								if (abs(CurrentSettlementAmount + 1000) > 1e-6)
								{
									string s="";   //沉降量字符串
									SaveTwoDecimal(CurrentSettlementAmount, s);
									ws.cell(2, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentSettleSpeed + 1000) > 1e-6)
								{
									string s = "";   //沉降速率字符串
									SaveThreeDecimal(CurrentSettleSpeed, s);
									ws.cell(3, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
							}
							else if (CurrentPeriod == 2)
							{
								if (abs(CurrentSettlementAmount + 1000) > 1e-6)
								{
									string s = "";   //沉降量字符串
									SaveTwoDecimal(CurrentSettlementAmount, s);
									ws.cell(4, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentASettleAmount + 1000) > 1e-6)
								{
									string s = "";   //累计沉降量字符串
									SaveTwoDecimal(CurrentASettleAmount, s);
									ws.cell(5, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentSettleSpeed + 1000) > 1e-6)
								{
									string s = "";   //沉降速率字符串
									SaveThreeDecimal(CurrentSettleSpeed, s);
									ws.cell(6, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}

							}
							else if (CurrentPeriod == 3)
							{
								if (abs(CurrentSettlementAmount + 1000) > 1e-6)
								{
									string s = "";   //沉降量字符串
									SaveTwoDecimal(CurrentSettlementAmount, s);
									ws.cell(7, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentASettleAmount + 1000) > 1e-6)
								{
									string s = "";   //累计沉降量字符串
									SaveTwoDecimal(CurrentASettleAmount, s);
									ws.cell(8, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentSettleSpeed + 1000) > 1e-6)
								{
									string s = "";   //沉降速率字符串
									SaveThreeDecimal(CurrentSettleSpeed, s);
									ws.cell(9, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
							}
							else
							{
								if (abs(CurrentSettlementAmount + 1000) > 1e-6)
								{
									string s = "";   //沉降量字符串
									SaveTwoDecimal(CurrentSettlementAmount, s);
									ws.cell(2+f*3, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentASettleAmount + 1000) > 1e-6)
								{
									string s = "";   //累计沉降量字符串
									SaveTwoDecimal(CurrentASettleAmount, s);
									ws.cell(2 + f * 3+1, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentSettleSpeed + 1000) > 1e-6)
								{
									string s = "";   //沉降速率字符串
									SaveThreeDecimal(CurrentSettleSpeed, s);
									ws.cell(2 + f * 3 + 1+1, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
							}
							CurrentPeriod++;
							f++;
						}
						CurrentPeriod -= f;
						f = 0;
					}
					int f = 0;
					//写入平均值
					int AverageRowNum = StartRow + 1 + P.qy[i].ContainGZW[j].ContainSettlementPoint.size() + 1 + 1;    //平均值所在的行数
					ws.cell(1, AverageRowNum).value(String_To_UTF8("平均值"));
					for (int k = CurrentPeriod; k < min(limit, CurrentRegionPeriod); k++)
					{
						double AverageSettlementAmount = P.qy[i].ContainGZW[j].AverageSettlementAmount[k]; //平均沉降量
						double AverageASettlementAmount= P.qy[i].ContainGZW[j].AverageAccumulateSettlementAmount[k];  //平均累积沉降量
						double AverageSettlementSpeed = P.qy[i].ContainGZW[j].AverageSettlementRate[k];    //平均沉降速率
						string AverageSettlementAmountStr = "";
						string AverageASettlementAmountStr = "";
						string AverageSettlementSpeedStr = "";
						SaveTwoDecimal(AverageSettlementAmount, AverageSettlementAmountStr);
						SaveTwoDecimal(AverageASettlementAmount, AverageASettlementAmountStr);
						SaveThreeDecimal(AverageSettlementSpeed, AverageSettlementSpeedStr);
						if (k == 1)
						{
							if (abs(AverageSettlementAmount + 1000) > 1e-6)
							{
								ws.cell(2, AverageRowNum).value(String_To_UTF8(AverageSettlementAmountStr));
							}
							if (abs(AverageSettlementSpeed + 1000) > 1e-6)
							{
								ws.cell(3, AverageRowNum).value(String_To_UTF8(AverageSettlementSpeedStr));
							}
						}
						else if (k == 2)
						{

							if (abs(AverageSettlementAmount + 1000) > 1e-6)
							{
								ws.cell(4, AverageRowNum).value(String_To_UTF8(AverageSettlementAmountStr));
							}
							if (abs(AverageASettlementAmount + 1000) > 1e-6)
							{
								ws.cell(5, AverageRowNum).value(String_To_UTF8(AverageASettlementAmountStr));
							}
							if (abs(AverageSettlementSpeed + 1000) > 1e-6)
							{
								ws.cell(6, AverageRowNum).value(String_To_UTF8(AverageSettlementSpeedStr));
							}
						}
						else if (k == 3)
						{
							if (abs(AverageSettlementAmount + 1000) > 1e-6)
							{
								ws.cell(7, AverageRowNum).value(String_To_UTF8(AverageSettlementAmountStr));
							}
							if (abs(AverageASettlementAmount + 1000) > 1e-6)
							{
								ws.cell(8, AverageRowNum).value(String_To_UTF8(AverageASettlementAmountStr));
							}
							if (abs(AverageSettlementSpeed + 1000) > 1e-6)
							{
								ws.cell(9, AverageRowNum).value(String_To_UTF8(AverageSettlementSpeedStr));
							}
						}
						else
						{
							if (abs(AverageSettlementAmount + 1000) > 1e-6)
							{
								ws.cell(2+f*3, AverageRowNum).value(String_To_UTF8(AverageSettlementAmountStr));
							}
							if (abs(AverageASettlementAmount + 1000) > 1e-6)
							{
								ws.cell(2 + f * 3+1, AverageRowNum).value(String_To_UTF8(AverageASettlementAmountStr));
							}
							if (abs(AverageSettlementSpeed + 1000) > 1e-6)
							{
								ws.cell(2 + f * 3+2, AverageRowNum).value(String_To_UTF8(AverageSettlementSpeedStr));
							}
						}
					}
					//写入说明
					ws.cell(1, AverageRowNum+1).value(String_To_UTF8("说明"));
					ws.cell(2,AverageRowNum+1).value(String_To_UTF8("1．表中数据单位为mm；\n2．“+”表示上升，“-”表示下沉。"));
		
					StartRow = AverageRowNum + 3;
					CurrentPeriod = min(limit, CurrentRegionPeriod);
				}
			}
		}
		wb.remove_sheet(sheet1);
		wb.save(path);
	}
	//生成分析结果
	void GenerateAnalysisResult( Project& P/*项目*/, const string& path/*保存路径*/)
	{
		std::ofstream outfile(path); // 创建ofstream对象，并打开文件

		if (!outfile.is_open()) { // 检查文件是否成功打开
			std::cerr << "无法打开文件" << std::endl;
			system("pause");
			exit(1);
		}
		//写入分析结果
		for (int i = 0; i < P.qy.size(); i++)
		{
			//写入区域名称
			outfile << RegionName[i]<< std::endl;


			for (int j = 0; j < P.qy[i].ContainGZW.size(); j++)
			{
				string CurrentGZWName = P.qy[i].ContainGZW[j].name;  //目前的构筑物名称
				int CurrentGZWAllPointNum = P.qy[i].ContainGZW[j].SettlementPointNum;  //目前的构筑物总点数
				int CurrentGZWUnstablePointNum = P.qy[i].ContainGZW[j].countUnstablePointNum();  //目前的构筑物不稳定的点数
				double CurrentMaxASettlementAmount = P.qy[i].ContainGZW[j].
					calcualteMaxAccumulateSettlementAmount();  //目前的构筑物最大累计沉降量
				double CurrentMinASettlementAmount = P.qy[i].ContainGZW[j].
					calcualteMinAccumulateSettlementAmount();  //目前的构筑物最小累计沉降量

				P.qy[i].ContainGZW[j].calcualteEachIssueAverageAccumulateSettlementAmount();
				double CurrentAverageASettlementAmount = P.qy[i].ContainGZW[j].
					AverageAccumulateSettlementAmount[P.qy[i].ContainGZW[j].frequency - 1]; //目前的构筑物平均累计沉降量
				double CurrentMaxSettlementSpeed = P.qy[i].ContainGZW[j].
					calcualteMaxSettlementRate();  //目前构筑物的最大沉降速率
				double CurrentMinSettlementSpeed = P.qy[i].ContainGZW[j].
					calcualteMinSettlementRate();  //目前构筑物的最小沉降速率 

				P.qy[i].ContainGZW[j].calculateEachIssueAverageSettlementSpeed();
				double CurrentAverageSettlementSpeed = P.qy[i].ContainGZW[j].AverageSettlementRate
					[P.qy[i].ContainGZW[j].frequency - 1]; //目前构筑物的平均沉降速率 

				//将构筑物的统计信息写入文本
				outfile << CurrentGZWName << std::endl;
				outfile << "总点数:" << CurrentGZWAllPointNum << endl;
				outfile << "未稳定点数:" << CurrentGZWUnstablePointNum << endl;
				if (abs(CurrentMaxASettlementAmount + 1000) <= 1e-6)
				{
					outfile << "最大累计沉降量:None" <<endl;
				}
				else
				{
					string s;
					outfile << "最大累计沉降量:" << fixed<<setprecision(2)<< 
						SaveTwoDecimal(CurrentMaxASettlementAmount,s)<<"mm"<< endl;
				}
				if (abs(CurrentMinASettlementAmount + 1000) <= 1e-6)
				{
					outfile << "最小累计沉降量:None" << endl;
				}
				else
				{
					string s;
					outfile << "最小累计沉降量:" << fixed << setprecision(2) <<
						SaveTwoDecimal(CurrentMinASettlementAmount, s) <<"mm" << endl;
				}
				if (abs(CurrentAverageASettlementAmount + 1000) <= 1e-6)
				{
					outfile << "平均累计沉降量:None" << endl;
				}
				else
				{
					string s;
					outfile << "平均累计沉降量:" << fixed << setprecision(2) <<
						SaveTwoDecimal(CurrentAverageASettlementAmount, s)<<"mm" << endl;
				}
				if (abs(CurrentMaxSettlementSpeed + 1000) <= 1e-6)
				{
					outfile << "最大沉降速率:None" << endl;
				}
				else
				{
					string s;
					outfile << "最大沉降速率:" << fixed << setprecision(3) <<
						SaveThreeDecimal(CurrentMaxSettlementSpeed, s)<<"mm/d" << endl;
				}
				if (abs(CurrentMinSettlementSpeed + 1000) <= 1e-6)
				{
					outfile << "最小沉降速率:None" << endl;
				}
				else
				{
					string s;
					outfile << "最小沉降速率:" << fixed << setprecision(3) <<
						SaveThreeDecimal(CurrentMinSettlementSpeed, s)<<"mm/d" << endl;
				}
				if (abs(CurrentAverageSettlementSpeed + 1000) <= 1e-6)
				{
					outfile << "平均沉降速率:None" << endl;
				}
				else
				{
					string s;
					outfile << "平均沉降速率:" << fixed << setprecision(3) <<
						SaveThreeDecimal(CurrentAverageSettlementSpeed, s)<<"mm/d" << endl;
				}
			}
			P.qy[i].calculateMinAverageAccumulateSettlementAmount();
			double MinAverageASettlementAmount = P.qy[i].
				MinAverageAccumulateSettlementAmount;  //该区域所有构筑物中最小的平均累计沉降量数值
			string MinAverageASettlementAmountGZWName = P.qy[i].
				MinAASettlementAmountGZW;   //该区域所有构筑物中最小的平均累计沉降量构筑物
			P.qy[i].calculateMaxAverageAccumulateSettlementAmount();
			double MaxAverageASettlementAmount = P.qy[i].
				MaxAverageAccumulateSettlementAmount;  //该区域所有构筑物中最大的平均累计沉降量数值
			string MaxAverageASettlementAmountGZWName = P.qy[i].
				MaxAASettlementAmountGZW;   //该区域所有构筑物中最大的平均累计沉降量构筑物 
			P.qy[i].calculateMinAverageSettlementRate();
			double MinAverageSettlementRate = P.qy[i].
				MinAverageSettlementRate;  //该区域所有构筑物中最小的平均沉降速率数值
			string MinAverageSettlementRateGZWName = P.qy[i].
				MinASettlementRateGZW;   //该区域所有构筑物中最小的平均沉降速率构筑物
			P.qy[i].calculateMaxAverageSettlementRate();
			double MaxAverageSettlementRate = P.qy[i].
				MaxAverageSettlementRate;  //该区域所有构筑物中最大的平均沉降速率数值
			string MaxAverageSettlementRateGZWName = P.qy[i].
				MaxASettlementRateGZW;   //该区域所有构筑物中最大的平均沉降速率构筑物
			P.qy[i].calculateMaxSettlementRate();
			double CurrentRegionMaxSettlementRate = P.qy[i].MaxSettlementRate;  //该区域所有沉降点中最大的沉降速率数值
			string CurrentRegionMaxSettlementRatePointName = P.qy[i].
				MaxSettlementRatePoint;   //该区域中最大的沉降速率沉降点
			string CurrentRegionMaxSettlementRateGZWName = P.qy[i].
				MaxSettlementRateGZW;  //该区域中最大的沉降速率沉降点所属的构筑物
			int UnStableGZWNum = P.qy[i].countMoreLimitGZW();   //该区域中未达到稳定标准的构筑物数


			outfile << endl;
			//将该区域的统计信息写入文本
			outfile << "最小的平均累计沉降量构筑物:" << MinAverageASettlementAmountGZWName << endl;
			string s;
			outfile << "最小的平均累计沉降量:" << fixed << setprecision(2) <<
				SaveTwoDecimal(MinAverageASettlementAmount, s) <<"mm"<< endl;
			
			outfile << "最大的平均累计沉降量构筑物:" << MaxAverageASettlementAmountGZWName << endl;
			outfile << "最大的平均累计沉降量:" << fixed << setprecision(2) <<
				SaveTwoDecimal(MaxAverageASettlementAmount, s) <<"mm"<< endl;
			
			outfile << "最小的平均沉降速率构筑物:" << MinAverageSettlementRateGZWName << endl;
			outfile << "最小的平均沉降速率:" << fixed << setprecision(3) <<
				SaveThreeDecimal(MinAverageSettlementRate, s)<<"mm/d" << endl;
			
			outfile << "最大的平均沉降速率构筑物:" << MaxAverageSettlementRateGZWName << endl;
			outfile << "最大的平均沉降速率:" << fixed << setprecision(3) <<
				SaveThreeDecimal(MaxAverageSettlementRate, s) << "mm/d" << endl;
			
			outfile << "最大的沉降速率构筑物:" << CurrentRegionMaxSettlementRateGZWName << endl;
			outfile << "最大的沉降速率点:" << CurrentRegionMaxSettlementRatePointName << endl;
			outfile << "最大的沉降速率:" << fixed << setprecision(3) <<
				SaveThreeDecimal(CurrentRegionMaxSettlementRate, s) << "mm/d" << endl;

			outfile << "未达到稳定标准的构筑物数:" << UnStableGZWNum << endl << endl << endl;
		}
		outfile.close(); // 关闭文件
	}
};


#endif 


