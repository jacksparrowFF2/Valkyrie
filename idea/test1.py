excode = {
	"日期":"20190830",
	"实验目的":"改变",
	"实验过程":"没有出现故障",
	"初始输入功率(w)":"81",
	"初始反馈功率(w)":"31",
	"末端输入功率(w)":"86",
	"末端反馈功率(w)":"36",
	"Ar(sccm)":"150",
	"H2(sccm)":"0",
	"CH4(sccm)":"9",
	"压强(pa)":"200",
	"温度(℃)":"600",
	"持续时间(min)":"60",
	"衬底1":"n-Si",
	"衬底2":"Quartz",
	"金属网":"MK1-0.5/0.5",
	"初步实验结果":"ssAA啊啊啊",
}
{
	"日期":"20190830",
	"实验目的":"改变",
	"实验过程":"没有出现故障",
	"初始输入功率(w)":"81",
	"初始反馈功率(w)":"31",
	"末端输入功率(w)":"86",
	"末端反馈功率(w)":"36",
	"Ar(sccm)":"150",
	"H2(sccm)":"0",
	"CH4(sccm)":"9",
	"压强(pa)":"200",
	"温度(℃)":"600",
	"持续时间(min)":"60",
	"衬底1":"n-Si",
	"衬底2":"Quartz",
	"金属网":"MK1-0.5/0.5",
	"初步实验结果":"ssAA啊啊啊",
}

date = int(excode["日期"])
power = int(excode["初始输入功率(w)"])-int(excode["初始反馈功率(w)"])
Ar = int(excode["Ar(sccm)"])
H2 = int(excode["H2(sccm)"])
CH4 = int(excode["CH4(sccm)"])
pressure = int(excode["CH4(sccm)"])
temp = int(excode["温度(℃)"])
sub1 = excode["衬底1"]
sub2 =excode["衬底2"]
metaltype = excode["金属网"]
note = excode["实验目的"]

# print(date,power,Ar,H2,CH4,pressure,temp,sub1+sub2,metaltype,note)
