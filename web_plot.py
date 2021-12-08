# -*- coding: utf-8 -*-
"""
Created on Fri Jun 14 14:47:10 2019

@author: msalam
"""

import matplotlib.pyplot as plt
#import time
import webbrowser
# import numpy as np

#webbrowser.open("file:///Users/drsalam/Desktop/Teaching/CMPS-599%20Project/Spring-2019/Rohit%20Varma%20Dandu/Code/Login_form.html")
#webbrowser.open('http://localhost/phppython/Login_form.html')

webbrowser.open('Login_form.html')

X1, Y1 = [], []
X2, Y2 = [], []
X3, Y3 = [], []
X4, Y4 = [], []
X5, Y5 = [], []

Xx1, Yy1 = [], []
Xx2, Yy2 = [], []
Xx3, Yy3 = [], []
Xx4, Yy4 = [], []
Xx5, Yy5 = [], []

#reading data from the file. This file has two columns.
#First column is the index val and second col is the data value
#X1 is getting index value and Y1 is getting data value
for line in open('611D81.txt', 'r'):
  values = [float(s) for s in line.split()]
  X1.append(values[0])
  Y1.append(values[1])
  
for line in open('611d87.txt', 'r'):
  values = [float(s) for s in line.split()]
  X2.append(values[0])
  Y2.append(values[1])
  
for line in open('611db8.txt', 'r'):
  values = [float(s) for s in line.split()]
  X3.append(values[0])
  Y3.append(values[1])
  
for line in open('612340.txt', 'r'):
  values = [float(s) for s in line.split()]
  X4.append(values[0])
  Y4.append(values[1])
  
for line in open('612405.txt', 'r'):
  values = [float(s) for s in line.split()]
  X5.append(values[0])
  Y5.append(values[1])

# getting the lengths/sizes of the data files

lenx = [len(X1),len(X2),len(X3),len(X4),len(X5)]

#print(min(lenx))

#copying the index and data from X1 and Y1 to Xx1 and Yy1, respectively
# up to the minimum index values


for i in range(min(lenx)):
  Xx1.append(X1[i])
  Yy1.append(Y1[i])
  Xx2.append(X2[i])
  Yy2.append(Y2[i])
  Xx3.append(X3[i])
  Yy3.append(Y3[i])
  Xx4.append(X4[i])
  Yy4.append(Y4[i])
  Xx5.append(X5[i])
  Yy5.append(Y5[i])
 

  if len(Xx1)==50: # this parameter gives the number of values to be display at a time
    #Line gragh plot
    plt.plot(Xx1, Yy1)
    plt.xlabel('Time')
    plt.ylabel('Sensor response')
    plt.title('Sensor-1 response System')
    plt.savefig('foo1.png')
    plt.close()
    #Bar gragh plot
    plt.bar(Xx1, Yy1, align='center', alpha=0.3)
    #y_pos = np.arange(len(Xx1))
    # Create names on the x-axis
    #plt.xticks(y_pos, Xx1)
    plt.title('Sensor-1 response')
    plt.xlabel('Time')
    plt.ylabel('Sensor data')
    plt.ylim([460,480])
    plt.savefig('foo1bar.png')
    plt.close()
    plt.pause(0.3)
    
    #Line gragh plot
    plt.plot(Xx2, Yy2)
    plt.xlabel('Time')
    plt.ylabel('Sensor response')
    plt.title('Sensor-2 response System')
    plt.savefig('foo2.png')
    plt.close()
    #Bar gragh plot
    plt.bar(Xx2, Yy2, align='center', alpha=0.5)
    plt.ylim([460,480])
    plt.title('Sensor-2 response')
    plt.xlabel('Time')
    plt.ylabel('Sensor data')
    plt.savefig('foo2bar.png')
    plt.close()
    plt.pause(0.3)
    
    #Line gragh plot
    plt.plot(Xx3, Yy3)
    plt.xlabel('Time')
    plt.ylabel('Sensor response')
    plt.title('Sensor-3 response System')
    plt.savefig('foo3.png')
    plt.close()
    #Bar gragh plot
    plt.bar(Xx3, Yy3, align='center', alpha=0.5)
    plt.ylim([460,480])
    plt.title('Sensor-3 response')
    plt.xlabel('Time')
    plt.ylabel('Sensor data')
    plt.savefig('foo3bar.png')
    plt.close()
    plt.pause(0.3)
    
    #Line gragh plot
    plt.plot(Xx4, Yy4)
    plt.xlabel('Time')
    plt.ylabel('Sensor response')
    plt.title('Sensor-4 response System')
    plt.savefig('foo4.png')
    plt.close()
    #Bar gragh plot
    plt.bar(Xx4, Yy4, align='center', alpha=0.5)
    plt.title('Sensor-4 response')
    plt.xlabel('Time')
    plt.ylabel('Sensor data')
    plt.ylim([460,480])
    plt.savefig('foo4bar.png')
    plt.close()
    plt.pause(0.3)
    
    #Line gragh plot
    plt.plot(Xx5, Yy5)
    plt.xlabel('Time')
    plt.ylabel('Sensor response')
    plt.title('Sensor-5 response System')
    plt.savefig('foo5.png')
    plt.close()
    #Bar gragh plot
    plt.bar(Xx5, Yy5, align='center', alpha=0.5)
    plt.title('Sensor-5 response')
    plt.xlabel('Time')
    plt.ylabel('Sensor data')
    plt.ylim([460,480])
    plt.savefig('foo5bar.png')
    plt.close()
    plt.pause(3)
   
  
    Xx1=[]
    Yy1=[]
    Xx2=[]
    Yy2=[]
    Xx3=[]
    Yy3=[]
    Xx4=[]
    Yy4=[]
    Xx5=[]
    Yy5=[]
   