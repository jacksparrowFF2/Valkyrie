{
	"日期":"数字",
	"实验目的":"text",
	"实验过程":"text",
	"初始输入功率(w)":"数字",
	"初始反馈功率(w)":"数字",
	"末端输入功率(w)":"数字",
	"末端反馈功率(w)":"数字",
	"Ar(sccm)":"数字",
	"H2(sccm)":"数字",
	"CH4(sccm)":"数字",
	"压强(pa)":"数字",
	"温度(℃)":"数字",
	"持续时间(min)":"数字",
	"衬底1":"text",
	"衬底2":"text",
	"金属网":"text",
	"初步实验结果":"text",
	"方阻(kΩ/□)":"数字",
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
