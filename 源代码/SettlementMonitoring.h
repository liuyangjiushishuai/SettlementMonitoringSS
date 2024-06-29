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
const string KnownPointName = "c02";    //��֪������
const double KnownNumberHeight = 33.83300;//��֪��߳�
const double SettlementRateLimit = 0.04;  //���������޲�
const vector<string>RegionName = { "���Ƹ��¸����ִ����޹�˾��̨��ҵ��Һ�廯��Ʒ��������׮��������������۲�ɹ���",
"���Ƹ��¸����ִ����޹�˾��̨��ҵ��Һ�廯��Ʒ�������̷�׮��������������۲�ɹ���" };    //��������
const string ProjectName = "���Ƹ��¸����ִ����޹�˾��̨��ҵ��Һ�廯��Ʒ�������̳����۲�";  //��������
const string CompanyName = "���Ƹ۸ۿڹ�������о�Ժ���޹�˾";                          //��˾����
const string NewestMeasureData = "";                //����һ�ڵĲ�������
const vector<string>NewAddPoint = {};   //�¼ӵĵ�
const vector<string>NewAddPointBelongGZW = {};  //�¼ӵĵ������Ĺ�����
const vector<int>NewAddPointBelongQY = {};    //�¼ӵĵ�����������
const vector<string>TableOrder = {" ","һ","��","��","��","��","��","��","��","��",
"ʮ","ʮһ","ʮ��","ʮ��","ʮ��","ʮ��","ʮ��","ʮ��","ʮ��","ʮ��","��ʮ"};        //����

//�����ڻ�Ϊ�ꡢ�¡���
vector<int>splitData(const string data);

//����ʱ����
int calculateTimeInterval(const vector<int>f, const vector<int>b);

//UTF-8�����ʽ�ַ��� ת��ͨsting����
std::string UTF8_To_string(const std::string& str);


//Unicode ת utf-8
bool WStringToString(const std::wstring& wstr, std::string& str);

//��ͨ�ַ���ת���ַ���
std::wstring to_wstring(const std::string& str);

//��ͨ�ַ���תUTF-8�ַ���
std::string String_To_UTF8(const std::string& str);

//������λС��
double SaveThreeDecimal(double OriginalData,string&str);

//������λС��
double SaveTwoDecimal(double OriginalData,string&str);
//��վ
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
//������
class cj_point
{
public:
	vector<string> cl_data; //��������
	vector<double> cl_height; //�����߳�
	int Cl_Frequency;    //��������
	vector<double> SettlementAmount;  //������
	vector<double> AccumulateSettlementAmount;  //�ۻ�������
	vector<double> SettlementSpeed;  //�����ٶ�
	string name;     //�����������

	cj_point(int n, string name)
	{
		this->Cl_Frequency = n;
		this->name = name;
	}
	//���������
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
	//�����ۻ�������
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

	//��������ٶ�
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
						//����ʱ����
						double interval_time = calculateTimeInterval(f_time_result, b_time_result);
						//��������ٶ�
						res.push_back(SettlementAmount[i] / interval_time);
					}
				}
			}
		}
		this->SettlementSpeed = res;
		return res;
	}
};
//������
class GZW
{
public:
	string name;    //����������
	int SettlementPointNum; //�����ĳ�������
	int UnstablePointNum;     //δ�ȶ��ĵ���
	int frequency;        //�ù������������
	double LatestMaxAccumulateSettlementAmount;      //���һ������ۻ�������
	double LatestMinAccumulateSettlementAmount;      //���һ����С�ۻ�������
	double LatestMaxSettlementRate;                //���һ���������ٶ�
	double LatestMinSettlementRate;               //���һ����С�����ٶ�
	vector<cj_point>ContainSettlementPoint; //�����ĳ�����
	vector<double>AverageSettlementAmount; //ÿ�ڵ�ƽ��������
	vector<double>AverageSettlementRate;   //ÿ�ڵ�ƽ�������ٶ�
	vector<double>AverageAccumulateSettlementAmount; //ÿ�ڵ�ƽ���ۻ�������


	GZW(string n/*����������*/, int f/*��������*/)
	{
		this->name = n;
		this->frequency = f;
	}
	//����ÿ�ڵ�ƽ��������
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
	//����ÿ�ڵ�ƽ�������ٶ�
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
	//����ÿ�ڵ�ƽ���ۻ�������
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
	//�������һ������ۻ�������
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
			this->LatestMaxAccumulateSettlementAmount = res;      //���һ������ۻ�������
		}
		return res;
	}
	//�������һ�ڵ���С�ۻ�������
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
			this->LatestMinAccumulateSettlementAmount = res;      //���һ����С�ۻ�������
		}
		return res;
	}
	//�������һ�ڵ��������ٶ�
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
			this->LatestMaxSettlementRate = -1000;                //���һ���������ٶ�
		}
		else
		{
			this->LatestMaxSettlementRate = res;                //���һ���������ٶ�
		}
		return res;
	}
	//�������һ�ڵ���С�����ٶ�
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
	//ͳ��δ�ﵽ�ȶ���׼�ĵ���
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
//����
class QY
{
public:
	string name;  //��������
	string EngineeringName;  //��������
	string CompanyName;       //��˾����
	int frequency;      //��������������
	vector<GZW>ContainGZW;//�����Ĺ�����
	vector<string>data;   //���������������
	double MaxAverageAccumulateSettlementAmount; //���ƽ���ۼƳ�����
	double MinAverageAccumulateSettlementAmount; //��Сƽ���ۼƳ�����
	string MaxAASettlementAmountGZW;     //���ƽ���ۼƳ�����������
	string MinAASettlementAmountGZW;    //��Сƽ���ۼƳ�����������
	double MaxAverageSettlementRate; //���ƽ�������ٶ�
	double MinAverageSettlementRate; //��Сƽ�������ٶ�
	string MaxASettlementRateGZW;     //���ƽ�������ٶȹ�����
	string MinASettlementRateGZW;    //��Сƽ�������ٶȹ�����
	string MaxSettlementRatePoint;   //�������ٶȵ�
	double MaxSettlementRate;     //�������ٶ�
	string MaxSettlementRateGZW;   //�������ٶȹ�����
	int MoreLimitGZWNum;        //���޵Ĺ���������

	QY(string n/*��������*/, string EN/*��������*/, string CN/*��˾����*/,int frequency/*��������������*/)
	{
		this->name = n;
		this->EngineeringName = EN;
		this->CompanyName = CN;
		this->frequency = frequency;
	}

	//��������ƽ���ۻ�������
	void calculateMaxAverageAccumulateSettlementAmount()
	{
		double CompareValue = 0;
		double res = 0;
		int count = 0;
		string s;  //���ƽ���ۼƳ���������Ӧ�Ĺ�����
		for (int i = 0; i < ContainGZW.size(); i++)
		{
			double rate = ContainGZW[i].frequency;   //�ù���������Ƶ��
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
	//������С��ƽ���ۼƳ�����
	void calculateMinAverageAccumulateSettlementAmount()
	{
		double res = 1000000;
		double CompareValue = 1000000;
		string s;  //��Сƽ���ۼƳ���������Ӧ�Ĺ�����
		int count = 0;
		for (int i = 0; i < ContainGZW.size(); i++)
		{
			double rate = ContainGZW[i].frequency;   //�ù���������Ƶ��
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
	//�������ƽ�������ٶ�
	void calculateMaxAverageSettlementRate()
	{
		double res = 0;
		double CompareValue = 0;
		int count = 0;
		string s;  //���ƽ�������ٶ�����Ӧ�Ĺ�����
		for (int i = 0; i < ContainGZW.size(); i++)
		{
			double rate = ContainGZW[i].frequency;   //�ù���������Ƶ��
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
	//������Сƽ�������ٶ�
	void calculateMinAverageSettlementRate()
	{
		double res = 1000000;
		double CompareValue = 1000000;
		string s;  //��Сƽ�������ٶ�����Ӧ�Ĺ�����
		int count = 0;
		for (int i = 0; i < ContainGZW.size(); i++)
		{
			double rate = ContainGZW[i].frequency;   //�ù���������Ƶ��
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
	//�����������ٶ�
	void calculateMaxSettlementRate()
	{
		double res = 0;
		double CompareValue = 0;
		string s;  //�������ٶȵ�
		string s2; //�������ٶȹ�����
		int count = 0;
		for (int i = 0; i < ContainGZW.size(); i++)
		{
			int ContainPointNum = ContainGZW[i].ContainSettlementPoint.size(); //�ù������������ĵ���
			int frequency = ContainGZW[i].frequency;    //�ù��������������
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
	//ͳ�Ƴ��޵Ĺ�����
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
//��Ŀ
class Project
{
public:
	vector<QY>qy; //������������
	int RegionNum;  //��������
	vector<unordered_set<string>>ContainGZW;  //�������Ĺ�����
	vector<unordered_map<string, unordered_set<string>>>ContainCJP;//�������ĳ�����
	vector<unordered_map<string, vector<double>>>ContainHeight;  //�������������ĸ߳�
	vector<vector<string>>GZW_order;      //��˳��װ�Ĺ�����
	vector<unordered_map<string, vector<string>>>CJP_order;  //��˳��װ�ĳ�����



	Project(int RN)
	{
		this->RegionNum = RN;
		/*this->ContainGZW.resize(this->RegionNum);
		this->ContainCJP.resize(this->RegionNum);
		this->ContainHeight.resize(this->RegionNum);*/
	}
	//���س����������
	void LoadData(string path/*������������ļ�·��*/)
	{
		xlnt::workbook wb;
		wb.load(path);


		//����ѭ��
		for (int i = 0; i < RegionNum; i++)
		{
			this->ContainGZW.push_back(unordered_set<string>());
			this->ContainCJP.push_back(unordered_map<string, unordered_set<string>>());
			this->ContainHeight.push_back(unordered_map<string, vector<double>>());
			this->GZW_order.push_back(vector<string>());
			this->CJP_order.push_back(unordered_map<string, vector<string>>());
			vector<GZW>structures;  //������
			vector<string>time;
			auto ws = wb.sheet_by_index(i);
			int c1 = 0;
			int	StartQs = 0;   //��¼ÿ�����Ŀ�ʼ����
			int SheetNum = 0;   //��¼�����ı����
			int CurrentSheetTitleRow = 0;   //��¼Ŀǰ�ı��������
			string CurrentGZWName = "";     //Ŀǰ�Ĺ���������
			string CurrentCJPName = "";     //Ŀǰ�ĳ���������
			for (auto row : ws.rows(false))
			{
				int c2 = 0;
				int c3 = StartQs;  //ͳ�Ƹ̹߳۲����
				for (auto cell : row)
				{
					//����һ�У������������й۲����ڶ���
					if (c1 == 0 && cell.has_value())
					{
						time.push_back(UTF8_To_string(cell.to_string()));
					}
					//�����۲�������һ��
					else if (UTF8_To_string(cell.to_string()) == "�۲�����")
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
					//����������������һ��
					else if (UTF8_To_string(cell.to_string()) == "����������")
					{
						CurrentSheetTitleRow = c1;
						break;
					}
					//�洢������
					else if (cell.has_value() && UTF8_To_string(ws.cell(c2 + 1, CurrentSheetTitleRow + 1).to_string()) == "����������")
					{
						CurrentGZWName = UTF8_To_string(cell.to_string());    //Ŀǰ�Ĺ���������
						if (this->ContainGZW[i].count(CurrentGZWName) == 0)
						{
							this->ContainGZW[i].emplace(CurrentGZWName);
							this->ContainCJP[i].emplace(CurrentGZWName, unordered_set<string>());
							this->GZW_order[i].push_back(CurrentGZWName);
						}
					}
					//�洢������
					else if (cell.has_value() && UTF8_To_string(ws.cell(c2 + 1, CurrentSheetTitleRow + 1).to_string()) == "���")
					{
						CurrentCJPName = UTF8_To_string(cell.to_string()); //Ŀǰ�ĳ���������
						if (this->ContainCJP[i][CurrentGZWName].count(CurrentCJPName) == 0)
						{
							this->ContainCJP[i][CurrentGZWName].emplace(CurrentCJPName);
							this->ContainHeight[i].emplace(CurrentCJPName, vector<double>());
							this->ContainHeight[i][CurrentCJPName].resize(time.size(), -100);
							this->CJP_order[i][CurrentGZWName].push_back(CurrentCJPName);
						}
					}
					//��������̼߳���
					else if (UTF8_To_string(ws.cell(c2 + 1, CurrentSheetTitleRow + 1).to_string()).find("ֵ") != string::npos)
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
			//����������Ĺ�����
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
			//����һ������
			QY CurrentQY(RegionName[i], ProjectName, CompanyName,time.size());
			CurrentQY.ContainGZW = structures;
			CurrentQY.data = time;
			this->qy.push_back(CurrentQY);
		}
		//���
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
	//�Թ涨�ļ�����ָ��ַ���
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
	//��ȡ��վ��Ϣ
	void read_station_information(cz& station, const vector<string>station_line)
	{
		int flag = 0;
		for (int i = 0; i < station_line.size(); i++)
		{
			flag++;
			//�ָ��ַ���
			vector<string>s;
			splitString(s, '|', ' ', station_line[i]);
			//��һ�κ���
			if (flag == 1)
			{
				station.b_point = s[5];
				station.first_b_num = stod(s[9]);
				station.first_b_distance = stod(s[12]);
			}
			//��һ��ǰ��
			else if (flag == 2)
			{
				station.f_point = s[5];
				station.first_f_num = stod(s[9]);
				station.first_f_distance = stod(s[12]);
			}
			//�ڶ���ǰ��
			else if (flag == 3)
			{
				station.second_f_num = stod(s[9]);
				station.second_f_distance = stod(s[12]);
			}
			//�ڶ��κ���
			else if (flag == 4)
			{
				station.second_b_num = stod(s[9]);
				station.second_b_distance = stod(s[12]);
			}
		}
	}
	//��ȡˮ׼ԭʼ����
	void readFile(string path, vector<cz>& d)
	{
		//���ļ�
		std::ifstream file(path);
		std::string line;
		vector<string>data;  //��¼��������
		int flag = 0;//�ж��ǲ���һ����վ��
		int line_num = 0;//��¼��������
		vector<string>station_information; //��¼һ����վ������
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
					//����
					for (int i = 0; i < station_information.size(); i++)
					{
						cout << station_information[i] << endl;
					}
					//��ʼ����һ����վ����Ϣ
					cz station;
					read_station_information(station, station_information);
					d.push_back(station);  //����վװ��
					flag = 0;    //����վ������Ϊ0
					station_information.clear(); //����¼��վ���ݵ������ÿ�
				}
				std::cout << line << std::endl;
			}
			////��ȡ�պϲ�
			//vector<string>SplitedInformation;
			//splitString(SplitedInformation, '|', ' ', data[line_num - 3]);
			//difference = stod(SplitedInformation[6]);
			file.close();
		}
		else
		{
			std::cout << "�޷���Ҫ��ȡ���ļ�" << std::endl;
		}
	}
	//����պϲ�
	double calculateClosureError(vector<cz>& d)
	{
		double res = 0;
		for (int i = 0; i < d.size(); i++)
		{
			res += d[i].calculateUnDistributionDifference();
		}
		return res;
	}
	//��·�߳��ȷ���պϲ�
	void allocationClosureError(const string path, vector<cz>& d)
	{
		//����ˮ׼·���ܳ���
		double all_length = 0;
		double AFDistance = 0;
		double ABDistance = 0;
		//����պϲ�
		double ClosureError = calculateClosureError(d);
		cout << "��վ����:" << d.size() << endl;
		//�����ۻ����Ӿ���
		for (int i = 0; i < d.size(); i++)
		{
			ABDistance += d[i].calculate_b_distance();
		}
		//�����ۻ�ǰ�Ӿ���
		for (int i = 0; i < d.size(); i++)
		{
			AFDistance += d[i].calculate_f_distance();
		}
		cout << "�ۻ����Ӿ���:" << ABDistance << endl;
		cout << "�ۻ�ǰ�Ӿ���" << AFDistance << endl;
		for (int i = 0; i < d.size(); i++)
		{
			all_length += d[i].calculate_b_distance() + d[i].calculate_f_distance();
		}
		//����ÿǧ�ױպϲ�
		double ClosureError_km = -ClosureError / (all_length / 1000);
		//����ÿ����վ�ĸ߲������
		for (int i = 0; i < d.size(); i++)
		{
			d[i].calculate_v(ClosureError_km);
			//����ÿ���������ĸ߳�
			if (i == 0)
			{
				d[i].calculae_front_point_height(KnownNumberHeight);
			}
			else
			{
				d[i].calculae_front_point_height(d[i - 1].front_point_height);
			}
		}


		//��ÿ����ĸ߳�д���ı�
		std::ofstream outfile(path); // ����ofstream���󣬲����ļ�

		if (outfile.is_open()) // ����ļ��Ƿ�ɹ���	
		{
			outfile << d[0].b_point << "," << KnownNumberHeight << std::endl; // ����׼��д��
			//�������д��
			for (int i = 0; i < d.size(); i++)
			{
				outfile << d[i].f_point << "," << std::fixed << std::setprecision(3) << d[i].front_point_height << std::endl;
			}
			outfile.close(); // �ر��ļ�
		}
		else
		{
			std::cout << "�޷���Ҫд����ļ�" << std::endl;
		}
		cout << "�������" << endl;
	}
	//������һ�����ݼ�����Ŀ
	void addNewestData(vector<cz>& d, string path1/*ԭʼˮ׼����·��*/, string path2/*д��ˮ׼�����ļ�·��*/, int flag/*�ж���û���µ�,1���У�0��û��*/)
	{
		readFile(path1, d);
		allocationClosureError(path2, d);
		//�洢����һ�ڵĳ�����
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
		//������µ�
		if (flag)
		{
			if (NewAddPoint.size() == 0)
			{
				cout << "������¼ӵĵ�" << endl;
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
	//������ֱ���
	void calculateVariable()
	{
		//������������ۻ��������������ٶ�
		for (int i = 0; i < this->qy.size(); i++)
		{
			for (int j = 0; j < this->qy[i].ContainGZW.size(); j++)
			{
				for (int k = 0; k < this->qy[i].ContainGZW[j].ContainSettlementPoint.size(); k++)
				{
					//���������
					this->qy[i].ContainGZW[j].ContainSettlementPoint[k].calculateSettlementAmount();
					//�����ۻ�������
					this->qy[i].ContainGZW[j].ContainSettlementPoint[k].calculateASettlementAmount();
					//��������ٶ�
					this->qy[i].ContainGZW[j].ContainSettlementPoint[k].calculateSettlementSpeed();
				}
			}
		}
		/*����ÿ�ڵ�ƽ����������ƽ�������ٶȡ�ƽ���ۻ������������һ�ڵ������С�ۻ���������
		���һ�ڵ������С�����ٶȡ�ͳ�Ʋ��ȶ��ĳ�������*/
		for (int i = 0; i < this->qy.size(); i++)
		{
			for (int j = 0; j < this->qy[i].ContainGZW.size(); j++)
			{
				//����ÿ�ڵ�ƽ��������
				this->qy[i].ContainGZW[j].calculateEachIssueAverageSettlementAmount();
				//����ÿ�ڵ�ƽ���ۻ�������
				this->qy[i].ContainGZW[j].calcualteEachIssueAverageAccumulateSettlementAmount();
				//����ÿ�ڵ�ƽ�������ٶ�
				this->qy[i].ContainGZW[j].calculateEachIssueAverageSettlementSpeed();
				//�������һ�ڵ�����ۻ�������
				this->qy[i].ContainGZW[j].calcualteMaxAccumulateSettlementAmount();
				//�������һ�ڵ���С�ۻ�������
				this->qy[i].ContainGZW[j].calcualteMinAccumulateSettlementAmount();
				//�������һ�ڵ��������ٶ�
				this->qy[i].ContainGZW[j].calcualteMaxSettlementRate();
				//�������һ�ڵ���С�����ٶ�
				this->qy[i].ContainGZW[j].calcualteMinSettlementRate();
				//ͳ�Ʋ��ȶ��ĳ�������
				this->qy[i].ContainGZW[j].countUnstablePointNum();
			}
		}
		//�������ƽ���ۻ���������������Сƽ���ۻ����������������ƽ�������ٶȡ�������Сƽ�������ٶȡ������������ٶȡ�ͳ�Ƴ��޵Ĺ�����
		for (int i = 0; i < qy.size(); i++)
		{
			//�������ƽ���ۻ�������
			qy[i].calculateMaxAverageAccumulateSettlementAmount();
			//������Сƽ���ۻ�������
			qy[i].calculateMinAverageAccumulateSettlementAmount();
			//�������ƽ�������ٶ�
			qy[i].calculateMaxAverageSettlementRate();
			//������Сƽ�������ٶ�
			qy[i].calculateMinAverageSettlementRate();
			//�����������ٶ�
			qy[i].calculateMaxSettlementRate();
			//ͳ�Ƴ��޵Ĺ�����
			qy[i].countMoreLimitGZW();
		}
		//����
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
		//����
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
		//����
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
		//����
		/*for (int i = 0; i < this->qy.size(); i++)
		{
			cout << i + 1 << "    " << "���ƽ���ۻ�������:" << qy[i].MaxAASettlementAmountGZW<<","<<fixed << setprecision(2) <<
				SaveTwoDecimal(qy[i].MaxAverageAccumulateSettlementAmount)
				<< "��Сƽ���ۻ�������:" <<qy[i].MinAASettlementAmountGZW<<"," << SaveTwoDecimal(qy[i].MinAverageAccumulateSettlementAmount)
				<< "���ƽ�������ٶ�:" <<qy[i].MaxASettlementRateGZW<<","<< fixed << setprecision(3)
				<< SaveThreeDecimal(qy[i].MaxAverageSettlementRate) << "��Сƽ�������ٶ�:" <<qy[i].MinASettlementRateGZW<<","<<
				SaveThreeDecimal(qy[i].MinAverageSettlementRate) <<
				"�������ٶ�:" <<qy[i].MaxSettlementRateGZW<<","<<qy[i].MaxSettlementRatePoint<<","<<SaveThreeDecimal(qy[i].MaxSettlementRate) 
				<< "���޹���������:" << qy[i].MoreLimitGZWNum << endl;
		}*/
		//����
		for (int i = 0; i < this->qy.size(); i++)
		{
			for (int j = 0; j < this->qy[i].ContainGZW.size(); j++)
			{
				cout << this->qy[i].ContainGZW[j].name << "    " << "�ܵ���:" << this->qy[i].ContainGZW[j].
					SettlementPointNum << "  " << "���޵ĵ���:" << this->qy[i].ContainGZW[j].UnstablePointNum << endl;
			}
		}
	}
	//����������
	pair<int, int> calculateIntervalMonths()
	{

	}
};

//���ɳɹ�
class GenerateResults
{
public:
	
	GenerateResults()
	{

	}
	//���ɳɹ���
	void GenerateResultsTable(const Project&P/*��Ŀ*/,const string& path/*����·��*/)
	{
		xlnt::workbook wb;
		auto sheet1 = wb.active_sheet();
		sheet1.title("sheet0");
		for (int i = 0; i < P.qy.size(); i++)
		{
			string CurrentRegionName = RegionName[i]; //Ŀǰ����������
			auto ws = wb.copy_sheet(sheet1);
			int frequency = P.qy[i].frequency;   //���������������
			vector<string>data = P.qy[i].data;
			int c = 0;     //��¼����
			int StartFrequency = 0;  //��¼ÿ����Ŀ�ʼ����
			int CurrentFrequency = 0;   //��¼��ǰ������
			int StartRowNum = 0;  //��¼ÿ����Ŀ�ʼ����
			while (CurrentFrequency<frequency)
			{
				c++;
				string CurrentTableName= CurrentRegionName +"��" + TableOrder[c] + "��"; //Ŀǰ�ı���
				int limit = 0;
				if (c == 1)
				{
					limit = CurrentFrequency + 3;
					//������д��
					ws.cell(1, StartRowNum+1).value(String_To_UTF8(CurrentTableName));
					//����������д��
					ws.cell(3, StartRowNum + 2).value(String_To_UTF8(ProjectName));
					ws.cell(1, StartRowNum + 2).value(String_To_UTF8("��������"));
					//���۲�����д��
					ws.cell(1, StartRowNum + 4).value(String_To_UTF8("�۲�����"));
					while (CurrentFrequency<min(limit,frequency))
					{
						if (CurrentFrequency == 0)
						{
							//����һ������д��
							ws.cell(3, StartRowNum + 4).value(String_To_UTF8(data[CurrentFrequency]));
						}
						else if (CurrentFrequency == 1)
						{
							//���ڶ��ڵ�����д��
							ws.cell(4, StartRowNum + 4).value(String_To_UTF8(data[CurrentFrequency]));
						}
						else if (CurrentFrequency == 2)
						{
							//�������ڵ�����д��
							ws.cell(7, StartRowNum + 4).value(String_To_UTF8(data[CurrentFrequency]));
						}
						CurrentFrequency++;
					}
					//д�빹��������
					ws.cell(1, StartRowNum + 6).value(String_To_UTF8("����������"));
					//д����
					ws.cell(2, StartRowNum + 6).value(String_To_UTF8("���"));
					//д���ͷ
					for (int j = StartFrequency; j < min(limit,frequency); j++)
					{
						string TableHeader = "��" + TableOrder[j + 1] + "�ι۲�ֵ";
						if (j == 0)
						{
							ws.cell(3, StartRowNum + 6).value(String_To_UTF8(TableHeader));
							ws.cell(3, StartRowNum + 7).value("(mm)");
						}
						else if (j == 1)
						{
							ws.cell(4, StartRowNum + 6).value(String_To_UTF8(TableHeader));
							ws.cell(4, StartRowNum + 7).value("(mm)");
							ws.cell(5, StartRowNum + 6).value(String_To_UTF8("���γ�����"));
							ws.cell(5, StartRowNum + 7).value("(mm)");
							ws.cell(6, StartRowNum + 6).value(String_To_UTF8("��������"));
							ws.cell(6, StartRowNum + 7).value("(mm/d)");
						}
						else if (j == 2)
						{
							ws.cell(7, StartRowNum + 6).value(String_To_UTF8(TableHeader));
							ws.cell(7, StartRowNum + 7).value("(mm)");
							ws.cell(8, StartRowNum + 6).value(String_To_UTF8("���γ�����"));
							ws.cell(8, StartRowNum + 7).value("(mm)");
							ws.cell(9, StartRowNum + 6).value(String_To_UTF8("�ۼƳ�����"));
							ws.cell(9, StartRowNum + 7).value("(mm)");
							ws.cell(10, StartRowNum + 6).value(String_To_UTF8("��������"));
							ws.cell(10, StartRowNum + 7).value("(mm/d)");
						}
					}
					//д������
					int GZWStartRow = StartRowNum + 7;   //��¼ÿ��������Ŀ�ʼ����
					for (int j = 0; j < P.qy[i].ContainGZW.size(); j++)
					{
						//�������������д��
						ws.cell(1, GZWStartRow + 1).value(String_To_UTF8(P.qy[i].ContainGZW[j].name));

						for (int m = 0; m < P.qy[i].ContainGZW[j].ContainSettlementPoint.size(); m++)
						{
							//������д��
							ws.cell(2, GZWStartRow+1 + m).value(String_To_UTF8(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].name));
							//������д��
							int f = 0;
							for (int k = StartFrequency; k < min(limit,frequency); k++)
							{
								if (k == 0)
								{
									if (abs(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k]+100)<=1e-6)
									{
										continue;
									}
									//���߳�����д��
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
									//���߳�����д��
									string NumStr = "";
									SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k], NumStr);
									ws.cell(3 + f, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
									//�����γ�����д��
									SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].SettlementAmount[k], NumStr);
									ws.cell(3 + f+1, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
									//����������д��
									SaveThreeDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].SettlementSpeed[k], NumStr);
									ws.cell(3 + f+2, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
								}
								else if (k == 2)
								{
									if (abs(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k] + 100) <= 1e-6)
									{
										continue;
									}
									//���߳�����д��
									string NumStr = "";
									SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k], NumStr);
									ws.cell(3 + f+2, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
									//�����γ�����д��
									SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].SettlementAmount[k], NumStr);
									ws.cell(3 + f + 3, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
									//���ۼƳ�����д��
									SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].AccumulateSettlementAmount[k], NumStr);
									ws.cell(3 + f + 4, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
									//����������д��
									SaveThreeDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].SettlementSpeed[k], NumStr);
									ws.cell(3 + f + 5, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
								}
								f++;
							}
						}
						GZWStartRow += P.qy[i].ContainGZW[j].ContainSettlementPoint.size(); //ˢ�¹����￪ʼ����
					}
					//���۲�������д��
					ws.cell(1, GZWStartRow + 1).value(String_To_UTF8("�۲�������(d)"));
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
					//��˵��д��
					ws.cell(1, GZWStartRow + 2).value(String_To_UTF8("˵��"));
					ws.cell(3, GZWStartRow + 2).value(String_To_UTF8("1����-����ʾ�³�����+����ʾ������"));
					//����ⵥλд��
					string MonitoringUnit="��ⵥλ��";
					MonitoringUnit += CompanyName;
					ws.cell(1, GZWStartRow + 4).value(String_To_UTF8(MonitoringUnit));
					//����ÿ����Ŀ�ʼ����
					StartRowNum = GZWStartRow + 5;
					//����ÿ����Ŀ�ʼ����
					StartFrequency = CurrentFrequency;
				}
				else
				{
					limit = CurrentFrequency + 2;
					//������д��
					ws.cell(1, StartRowNum + 1).value(String_To_UTF8(CurrentTableName));
					//����������д��
					ws.cell(3, StartRowNum + 2).value(String_To_UTF8(ProjectName));
					ws.cell(1, StartRowNum + 2).value(String_To_UTF8("��������"));
					//���۲�����д��
					ws.cell(1, StartRowNum + 4).value(String_To_UTF8("�۲�����"));
					int f = 0;
					while (CurrentFrequency < min(limit, frequency))
					{
						ws.cell(3+4*f, StartRowNum + 4).value(String_To_UTF8(data[CurrentFrequency]));
						CurrentFrequency++;
						f++;
					}
					f = 0;
					//д�빹��������
					ws.cell(1, StartRowNum + 6).value(String_To_UTF8("����������"));
					//д����
					ws.cell(2, StartRowNum + 6).value(String_To_UTF8("���"));
					//д���ͷ
					for (int j = StartFrequency; j < min(limit, frequency); j++)
					{
						string TableHeader = "��" + TableOrder[j + 1] + "�ι۲�ֵ";
						ws.cell(3+4*f, StartRowNum + 6).value(String_To_UTF8(TableHeader));
						ws.cell(3+4*f, StartRowNum + 7).value("(mm)");
						ws.cell(3 + 4 * f+1, StartRowNum + 6).value(String_To_UTF8("���γ�����"));
						ws.cell(3 + 4 * f + 1, StartRowNum + 7).value("(mm)");
						ws.cell(3+4*f+2, StartRowNum + 6).value(String_To_UTF8("�ۼƳ�����"));
						ws.cell(3 + 4 * f + 2, StartRowNum + 7).value("(mm)");
						ws.cell(3+4*f+3, StartRowNum + 6).value(String_To_UTF8("��������"));
						ws.cell(3+4*f+3, StartRowNum + 7).value("(mm/d)");
						f++;
					}
					//д������
					int GZWStartRow = StartRowNum + 7;   //��¼ÿ��������Ŀ�ʼ����
					for (int j = 0; j < P.qy[i].ContainGZW.size(); j++)
					{
						//�������������д��
						ws.cell(1, GZWStartRow + 1).value(String_To_UTF8(P.qy[i].ContainGZW[j].name));

						for (int m = 0; m < P.qy[i].ContainGZW[j].ContainSettlementPoint.size(); m++)
						{
							//������д��
							ws.cell(2, GZWStartRow + 1 + m).value(String_To_UTF8(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].name));
							//������д��
							int f = 0;
							for (int k = StartFrequency; k < min(limit, frequency); k++)
							{
								if (abs(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k] + 100) <= 1e-6)
								{
									continue;
								}
								//���߳�����д��
								string NumStr = "";
								SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].cl_height[k], NumStr);
								ws.cell(3 + 4*f, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
								//�����γ�����д��
								SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].SettlementAmount[k], NumStr);
								ws.cell(3 + 4*f + 1, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
								//���ۼƳ�����д��
								SaveTwoDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].AccumulateSettlementAmount[k], NumStr);
								ws.cell(3 + 4*f + 2, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
								//����������д��
								SaveThreeDecimal(P.qy[i].ContainGZW[j].ContainSettlementPoint[m].SettlementSpeed[k], NumStr);
								ws.cell(3 + 4*f + 3, GZWStartRow + m+1).value(String_To_UTF8(NumStr));
								f++;
							}
						}
						GZWStartRow += P.qy[i].ContainGZW[j].ContainSettlementPoint.size(); //ˢ�¹����￪ʼ����
					}
					//���۲�������д��
					ws.cell(1, GZWStartRow + 1).value(String_To_UTF8("�۲�������(d)"));
					f = 0;
					for (int k = StartFrequency; k < min(limit, frequency); k++)
					{
						auto f_time = splitData(data[k - 1]);
						auto b_time = splitData(data[k]);
						int interval = calculateTimeInterval(f_time, b_time);
						ws.cell(3 + 4*f, GZWStartRow + 1).value(interval);
						f++;
					}
					//��˵��д��
					ws.cell(1, GZWStartRow + 2).value(String_To_UTF8("˵��"));
					ws.cell(3, GZWStartRow + 2).value(String_To_UTF8("1����-����ʾ�³�����+����ʾ������"));
					//����ⵥλд��
					string MonitoringUnit = "��ⵥλ��";
					MonitoringUnit += CompanyName;
					ws.cell(1, GZWStartRow + 4).value(String_To_UTF8(MonitoringUnit));
				}

			}
		}
		wb.remove_sheet(sheet1);
		wb.save(path);
	}
	//���ɳɹ�������
	void GenerateResultAnalysisTable(const Project& P/*��Ŀ*/, const string& path/*����·��*/)
	{
		xlnt::workbook wb;
		auto sheet1 = wb.active_sheet();
		sheet1.title("sheet0");
		for (int i = 0; i < P.qy.size(); i++)
		{
			auto ws = wb.copy_sheet(sheet1);
			int CurrentRegionPeriod = P.qy[i].frequency;
			vector<string>data = P.qy[i].data;   //Ŀǰ���������ʱ��
			int record = 0;
			int StartRow = 1;  //ÿ����Ŀ�ʼ����
			for (int j = 0; j < P.qy[i].ContainGZW.size(); j++)
			{
				int CurrentPeriod = 1;
				while (CurrentPeriod< CurrentRegionPeriod)
				{
					//д�빹��������
					string CurrentGZWName = P.qy[i].ContainGZW[j].name;
					ws.cell(1, StartRow - 1 + 1).value(String_To_UTF8(CurrentGZWName));
					if (CurrentPeriod == 1)
					{
						//д���ͷ
						ws.cell(1, StartRow + 1).value(String_To_UTF8("���"));
						ws.cell(2, StartRow + 1).value(String_To_UTF8(data[CurrentPeriod]));
						if (CurrentPeriod + 1 < CurrentRegionPeriod)
						{
							ws.cell(4, StartRow + 1).value(String_To_UTF8(data[CurrentPeriod+1]));
						}
						if (CurrentPeriod + 2 < CurrentRegionPeriod)
						{
							ws.cell(7, StartRow + 1).value(String_To_UTF8(data[CurrentPeriod + 2]));
						}
						ws.cell(2, StartRow + 2).value(String_To_UTF8("���γ�����(mm)"));
						ws.cell(3, StartRow + 2).value(String_To_UTF8("��������(mm/d)"));
						
						ws.cell(4, StartRow + 2).value(String_To_UTF8("���γ�����(mm)"));
						ws.cell(5, StartRow + 2).value(String_To_UTF8("�ۼƳ�����(mm)"));
						ws.cell(6, StartRow + 2).value(String_To_UTF8("��������(mm/d)"));
						
						ws.cell(7, StartRow + 2).value(String_To_UTF8("���γ�����(mm)"));
						ws.cell(8, StartRow + 2).value(String_To_UTF8("�ۼƳ�����(mm)"));
						ws.cell(9, StartRow + 2).value(String_To_UTF8("��������(mm/d)"));
					}
					else
					{
						ws.cell(1, StartRow + 1).value(String_To_UTF8("���"));
						int f = 0;
						for (int i = CurrentPeriod; i < CurrentPeriod + 3; i++)
						{
							if (i < CurrentRegionPeriod)
							{
								ws.cell(2 + f * 3, StartRow + 1).value(String_To_UTF8(data[i]));
							}
							ws.cell(2 + f * 3, StartRow + 2).value(String_To_UTF8("���γ�����(mm)"));
							ws.cell(2 + f * 3 + 1, StartRow + 2).value(String_To_UTF8("�ۼƳ�����(mm)"));
							ws.cell(2 + f * 3 + 1 + 1, StartRow + 2).value(String_To_UTF8("��������(mm/d)"));
							f++;
						}

					}
					int limit = CurrentPeriod + 3;
					for (int m = 0;m< P.qy[i].ContainGZW[j].ContainSettlementPoint.size(); m++)
					{
						//д�����
						string CurrentPointName = P.qy[i].ContainGZW[j].ContainSettlementPoint[m].name;
						ws.cell(1, StartRow + 2 + m + 1).value(String_To_UTF8(CurrentPointName));
						int f = 0;
						while (CurrentPeriod<min(limit, CurrentRegionPeriod))
						{
							double CurrentSettlementAmount = P.qy[i].ContainGZW[j].ContainSettlementPoint[m].
								SettlementAmount[CurrentPeriod];  //Ŀǰ�ĳ�����
							double CurrentASettleAmount = P.qy[i].ContainGZW[j].ContainSettlementPoint[m].
								AccumulateSettlementAmount[CurrentPeriod];    //Ŀǰ���ۼƳ�����
							double CurrentSettleSpeed = P.qy[i].ContainGZW[j].ContainSettlementPoint[m].
								SettlementSpeed[CurrentPeriod];  //Ŀǰ�ĳ����ٶ�
							if (CurrentPeriod == 1)
							{
								if (abs(CurrentSettlementAmount + 1000) > 1e-6)
								{
									string s="";   //�������ַ���
									SaveTwoDecimal(CurrentSettlementAmount, s);
									ws.cell(2, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentSettleSpeed + 1000) > 1e-6)
								{
									string s = "";   //���������ַ���
									SaveThreeDecimal(CurrentSettleSpeed, s);
									ws.cell(3, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
							}
							else if (CurrentPeriod == 2)
							{
								if (abs(CurrentSettlementAmount + 1000) > 1e-6)
								{
									string s = "";   //�������ַ���
									SaveTwoDecimal(CurrentSettlementAmount, s);
									ws.cell(4, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentASettleAmount + 1000) > 1e-6)
								{
									string s = "";   //�ۼƳ������ַ���
									SaveTwoDecimal(CurrentASettleAmount, s);
									ws.cell(5, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentSettleSpeed + 1000) > 1e-6)
								{
									string s = "";   //���������ַ���
									SaveThreeDecimal(CurrentSettleSpeed, s);
									ws.cell(6, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}

							}
							else if (CurrentPeriod == 3)
							{
								if (abs(CurrentSettlementAmount + 1000) > 1e-6)
								{
									string s = "";   //�������ַ���
									SaveTwoDecimal(CurrentSettlementAmount, s);
									ws.cell(7, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentASettleAmount + 1000) > 1e-6)
								{
									string s = "";   //�ۼƳ������ַ���
									SaveTwoDecimal(CurrentASettleAmount, s);
									ws.cell(8, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentSettleSpeed + 1000) > 1e-6)
								{
									string s = "";   //���������ַ���
									SaveThreeDecimal(CurrentSettleSpeed, s);
									ws.cell(9, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
							}
							else
							{
								if (abs(CurrentSettlementAmount + 1000) > 1e-6)
								{
									string s = "";   //�������ַ���
									SaveTwoDecimal(CurrentSettlementAmount, s);
									ws.cell(2+f*3, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentASettleAmount + 1000) > 1e-6)
								{
									string s = "";   //�ۼƳ������ַ���
									SaveTwoDecimal(CurrentASettleAmount, s);
									ws.cell(2 + f * 3+1, StartRow + 2 + m + 1).value(String_To_UTF8(s));
								}
								if (abs(CurrentSettleSpeed + 1000) > 1e-6)
								{
									string s = "";   //���������ַ���
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
					//д��ƽ��ֵ
					int AverageRowNum = StartRow + 1 + P.qy[i].ContainGZW[j].ContainSettlementPoint.size() + 1 + 1;    //ƽ��ֵ���ڵ�����
					ws.cell(1, AverageRowNum).value(String_To_UTF8("ƽ��ֵ"));
					for (int k = CurrentPeriod; k < min(limit, CurrentRegionPeriod); k++)
					{
						double AverageSettlementAmount = P.qy[i].ContainGZW[j].AverageSettlementAmount[k]; //ƽ��������
						double AverageASettlementAmount= P.qy[i].ContainGZW[j].AverageAccumulateSettlementAmount[k];  //ƽ���ۻ�������
						double AverageSettlementSpeed = P.qy[i].ContainGZW[j].AverageSettlementRate[k];    //ƽ����������
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
					//д��˵��
					ws.cell(1, AverageRowNum+1).value(String_To_UTF8("˵��"));
					ws.cell(2,AverageRowNum+1).value(String_To_UTF8("1���������ݵ�λΪmm��\n2����+����ʾ��������-����ʾ�³���"));
		
					StartRow = AverageRowNum + 3;
					CurrentPeriod = min(limit, CurrentRegionPeriod);
				}
			}
		}
		wb.remove_sheet(sheet1);
		wb.save(path);
	}
	//���ɷ������
	void GenerateAnalysisResult( Project& P/*��Ŀ*/, const string& path/*����·��*/)
	{
		std::ofstream outfile(path); // ����ofstream���󣬲����ļ�

		if (!outfile.is_open()) { // ����ļ��Ƿ�ɹ���
			std::cerr << "�޷����ļ�" << std::endl;
			system("pause");
			exit(1);
		}
		//д��������
		for (int i = 0; i < P.qy.size(); i++)
		{
			//д����������
			outfile << RegionName[i]<< std::endl;


			for (int j = 0; j < P.qy[i].ContainGZW.size(); j++)
			{
				string CurrentGZWName = P.qy[i].ContainGZW[j].name;  //Ŀǰ�Ĺ���������
				int CurrentGZWAllPointNum = P.qy[i].ContainGZW[j].SettlementPointNum;  //Ŀǰ�Ĺ������ܵ���
				int CurrentGZWUnstablePointNum = P.qy[i].ContainGZW[j].countUnstablePointNum();  //Ŀǰ�Ĺ����ﲻ�ȶ��ĵ���
				double CurrentMaxASettlementAmount = P.qy[i].ContainGZW[j].
					calcualteMaxAccumulateSettlementAmount();  //Ŀǰ�Ĺ���������ۼƳ�����
				double CurrentMinASettlementAmount = P.qy[i].ContainGZW[j].
					calcualteMinAccumulateSettlementAmount();  //Ŀǰ�Ĺ�������С�ۼƳ�����

				P.qy[i].ContainGZW[j].calcualteEachIssueAverageAccumulateSettlementAmount();
				double CurrentAverageASettlementAmount = P.qy[i].ContainGZW[j].
					AverageAccumulateSettlementAmount[P.qy[i].ContainGZW[j].frequency - 1]; //Ŀǰ�Ĺ�����ƽ���ۼƳ�����
				double CurrentMaxSettlementSpeed = P.qy[i].ContainGZW[j].
					calcualteMaxSettlementRate();  //Ŀǰ�����������������
				double CurrentMinSettlementSpeed = P.qy[i].ContainGZW[j].
					calcualteMinSettlementRate();  //Ŀǰ���������С�������� 

				P.qy[i].ContainGZW[j].calculateEachIssueAverageSettlementSpeed();
				double CurrentAverageSettlementSpeed = P.qy[i].ContainGZW[j].AverageSettlementRate
					[P.qy[i].ContainGZW[j].frequency - 1]; //Ŀǰ�������ƽ���������� 

				//���������ͳ����Ϣд���ı�
				outfile << CurrentGZWName << std::endl;
				outfile << "�ܵ���:" << CurrentGZWAllPointNum << endl;
				outfile << "δ�ȶ�����:" << CurrentGZWUnstablePointNum << endl;
				if (abs(CurrentMaxASettlementAmount + 1000) <= 1e-6)
				{
					outfile << "����ۼƳ�����:None" <<endl;
				}
				else
				{
					string s;
					outfile << "����ۼƳ�����:" << fixed<<setprecision(2)<< 
						SaveTwoDecimal(CurrentMaxASettlementAmount,s)<<"mm"<< endl;
				}
				if (abs(CurrentMinASettlementAmount + 1000) <= 1e-6)
				{
					outfile << "��С�ۼƳ�����:None" << endl;
				}
				else
				{
					string s;
					outfile << "��С�ۼƳ�����:" << fixed << setprecision(2) <<
						SaveTwoDecimal(CurrentMinASettlementAmount, s) <<"mm" << endl;
				}
				if (abs(CurrentAverageASettlementAmount + 1000) <= 1e-6)
				{
					outfile << "ƽ���ۼƳ�����:None" << endl;
				}
				else
				{
					string s;
					outfile << "ƽ���ۼƳ�����:" << fixed << setprecision(2) <<
						SaveTwoDecimal(CurrentAverageASettlementAmount, s)<<"mm" << endl;
				}
				if (abs(CurrentMaxSettlementSpeed + 1000) <= 1e-6)
				{
					outfile << "����������:None" << endl;
				}
				else
				{
					string s;
					outfile << "����������:" << fixed << setprecision(3) <<
						SaveThreeDecimal(CurrentMaxSettlementSpeed, s)<<"mm/d" << endl;
				}
				if (abs(CurrentMinSettlementSpeed + 1000) <= 1e-6)
				{
					outfile << "��С��������:None" << endl;
				}
				else
				{
					string s;
					outfile << "��С��������:" << fixed << setprecision(3) <<
						SaveThreeDecimal(CurrentMinSettlementSpeed, s)<<"mm/d" << endl;
				}
				if (abs(CurrentAverageSettlementSpeed + 1000) <= 1e-6)
				{
					outfile << "ƽ����������:None" << endl;
				}
				else
				{
					string s;
					outfile << "ƽ����������:" << fixed << setprecision(3) <<
						SaveThreeDecimal(CurrentAverageSettlementSpeed, s)<<"mm/d" << endl;
				}
			}
			P.qy[i].calculateMinAverageAccumulateSettlementAmount();
			double MinAverageASettlementAmount = P.qy[i].
				MinAverageAccumulateSettlementAmount;  //���������й���������С��ƽ���ۼƳ�������ֵ
			string MinAverageASettlementAmountGZWName = P.qy[i].
				MinAASettlementAmountGZW;   //���������й���������С��ƽ���ۼƳ�����������
			P.qy[i].calculateMaxAverageAccumulateSettlementAmount();
			double MaxAverageASettlementAmount = P.qy[i].
				MaxAverageAccumulateSettlementAmount;  //���������й�����������ƽ���ۼƳ�������ֵ
			string MaxAverageASettlementAmountGZWName = P.qy[i].
				MaxAASettlementAmountGZW;   //���������й�����������ƽ���ۼƳ����������� 
			P.qy[i].calculateMinAverageSettlementRate();
			double MinAverageSettlementRate = P.qy[i].
				MinAverageSettlementRate;  //���������й���������С��ƽ������������ֵ
			string MinAverageSettlementRateGZWName = P.qy[i].
				MinASettlementRateGZW;   //���������й���������С��ƽ���������ʹ�����
			P.qy[i].calculateMaxAverageSettlementRate();
			double MaxAverageSettlementRate = P.qy[i].
				MaxAverageSettlementRate;  //���������й�����������ƽ������������ֵ
			string MaxAverageSettlementRateGZWName = P.qy[i].
				MaxASettlementRateGZW;   //���������й�����������ƽ���������ʹ�����
			P.qy[i].calculateMaxSettlementRate();
			double CurrentRegionMaxSettlementRate = P.qy[i].MaxSettlementRate;  //���������г����������ĳ���������ֵ
			string CurrentRegionMaxSettlementRatePointName = P.qy[i].
				MaxSettlementRatePoint;   //�����������ĳ������ʳ�����
			string CurrentRegionMaxSettlementRateGZWName = P.qy[i].
				MaxSettlementRateGZW;  //�����������ĳ������ʳ����������Ĺ�����
			int UnStableGZWNum = P.qy[i].countMoreLimitGZW();   //��������δ�ﵽ�ȶ���׼�Ĺ�������


			outfile << endl;
			//���������ͳ����Ϣд���ı�
			outfile << "��С��ƽ���ۼƳ�����������:" << MinAverageASettlementAmountGZWName << endl;
			string s;
			outfile << "��С��ƽ���ۼƳ�����:" << fixed << setprecision(2) <<
				SaveTwoDecimal(MinAverageASettlementAmount, s) <<"mm"<< endl;
			
			outfile << "����ƽ���ۼƳ�����������:" << MaxAverageASettlementAmountGZWName << endl;
			outfile << "����ƽ���ۼƳ�����:" << fixed << setprecision(2) <<
				SaveTwoDecimal(MaxAverageASettlementAmount, s) <<"mm"<< endl;
			
			outfile << "��С��ƽ���������ʹ�����:" << MinAverageSettlementRateGZWName << endl;
			outfile << "��С��ƽ����������:" << fixed << setprecision(3) <<
				SaveThreeDecimal(MinAverageSettlementRate, s)<<"mm/d" << endl;
			
			outfile << "����ƽ���������ʹ�����:" << MaxAverageSettlementRateGZWName << endl;
			outfile << "����ƽ����������:" << fixed << setprecision(3) <<
				SaveThreeDecimal(MaxAverageSettlementRate, s) << "mm/d" << endl;
			
			outfile << "���ĳ������ʹ�����:" << CurrentRegionMaxSettlementRateGZWName << endl;
			outfile << "���ĳ������ʵ�:" << CurrentRegionMaxSettlementRatePointName << endl;
			outfile << "���ĳ�������:" << fixed << setprecision(3) <<
				SaveThreeDecimal(CurrentRegionMaxSettlementRate, s) << "mm/d" << endl;

			outfile << "δ�ﵽ�ȶ���׼�Ĺ�������:" << UnStableGZWNum << endl << endl << endl;
		}
		outfile.close(); // �ر��ļ�
	}
};


#endif 


