
1.python version 验证
输入:python --version
打印:Python 2.7.18

2.numpy
输入:pip install numpy
版本：1.16.6

3.pandas
输入:pip install pandas
版本: 0.24.2

4.xlrd
输入:pip install xlrd
版本: 2.0.1


5.xlwt
输入:pip install xlwt
版本:1.3.0

#使用说明
c表 Lij = Ecu_Dist, Lmax = E列最大值， i c列， j d列

p表 Pi（i板块p值）=PRIORITY_W, Si = Shape_Area(平方米转换平方千米)E列， Ci =LE。。。 H列 


Gij=Pi *Si * Pj * Sj * Lmax平方/ （Ci*Cj*Lij 平方）

Lij c表 e列， i= c列的值，j =D列的值

思路：c表中算每一行一个值，附加到最后。i取c列的值，j取D列的值。拿到i j后，去p表找G列对应的i j，然后取p表c列的值为Pi Pj，e列的值为Si Sj。
Lmax的平方是E列最大值。Ci H列， Cj H列
