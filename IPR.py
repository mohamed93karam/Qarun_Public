
import pyodbc
import datetime
import numpy as np

import matplotlib.pyplot as plt
import math, copy
from decimal import *
import re


# Import module
from tkinter import *






def main(completion):
    # completion = input("Enter Completion Name\n")
    latest_points = 20 # input("Enter latest number of points to use, default 20\n")
    latest_points_num = int(latest_points if latest_points else 20)


    conn = pyodbc.connect('Driver={SQL Server};'
        'Server=QPCAVODB2K12;'
        'Database=AVOCET_PRODUCTION;'
        'Trusted_Connection=yes;'
        'UID=sa;'
        'PWD=AAAAAAA;')

    cursor = conn.cursor()
    cursor.execute(f'''
    DECLARE @COMPLETION VARCHAR(MAX) = '{completion}';
    SELECT START_DATETIME, OIL_VOL, WATER_VOL, LIQ_VOL, BSW_VOL_FRAC, TEST_TYPE FROM VT_WELL_TEST_en_US WHERE
    ITEM_NAME = @COMPLETION
    AND
    START_DATETIME - 10 >= (SELECT MAX(START_DATETIME) FROM VI_COMPLETION_ALL_en_US WHERE ITEM_NAME = @COMPLETION AND [STATUS] = 'PRODUCING' AND TYPE = 'PRODUCTION')
    AND VALID_TEST = 'True'
    ''')

    all_well_test_data = list(cursor.fetchall())
    well_test_data = list(filter(lambda test: test[5] != 'VALID_ANALYSIS', all_well_test_data))[-(latest_points_num):]


    cursor.execute(f'''
    DECLARE @COMPLETION VARCHAR(MAX) = '{completion}';
    SELECT START_DATETIME, OIL_VOL, WATER_VOL, LIQ_VOL, BSW_VOL_FRAC, TEST_TYPE FROM VT_WELL_TEST_en_US WHERE
    ITEM_NAME = @COMPLETION
    AND VALID_TEST = 'True'
    ''')

    all_well_test_data = list(cursor.fetchall())



    cursor.execute(f'''
        DECLARE @COMPLETION VARCHAR(MAX) = '{completion}';
        SELECT DFL.START_DATETIME, DFL.FLUID_LEVEL, DFL.NLAP, DFL.SFL, DFL.COMMENT
        FROM VT_DFL_TEST_en_US DFL
        WHERE
        DFL.ITEM_NAME = @COMPLETION
        AND
        DFL.START_DATETIME >= (SELECT MAX(START_DATETIME) FROM VI_COMPLETION_ALL_en_US WHERE ITEM_NAME = @COMPLETION AND [STATUS] = 'PRODUCING' AND TYPE = 'PRODUCTION');
    ''')

    fluid_level_data = list(cursor.fetchall())



    cursor.execute(f'''
        DECLARE @COMPLETION VARCHAR(MAX) = '{completion}';

        SELECT MID_PERF from VI_COMPLETION_ALL_en_US
        WHERE
        ITEM_NAME = @COMPLETION
        AND
        START_DATETIME >= (SELECT MAX(START_DATETIME) FROM VI_COMPLETION_ALL_en_US WHERE ITEM_NAME = @COMPLETION AND [STATUS] = 'PRODUCING' AND TYPE = 'PRODUCTION')
    ''')

    mid_perf = cursor.fetchone()[0]

    
    cursor.execute(f'''
        DECLARE @COMPLETION VARCHAR(MAX) = '{completion}';

        SELECT DFL.START_DATETIME, coalesce(DFL.FLUID_LEVEL, DFL.SFL),  DFL.COMMENT, coalesce(DFL.FLUID_LEVEL, DFL.SFL), C.[STATUS]
        FROM VT_DFL_TEST_en_US DFL
        LEFT JOIN VI_COMPLETION_ALL_en_US C
        ON
        C.ITEM_ID = DFL.ITEM_ID
        AND
        C.START_DATETIME <= DFL.START_DATETIME
        AND
        C.END_DATETIME > DFL.START_DATETIME
        WHERE
        DFL.ITEM_NAME = @COMPLETION
        AND coalesce(DFL.FLUID_LEVEL, DFL.SFL) IS NOT NULL
        AND
        DFL.START_DATETIME >= (SELECT MAX(START_DATETIME) FROM VL_WELL_ZONE_ALL_en_US WHERE VL_WELL_ZONE_ALL_en_US.Completion = @COMPLETION AND [END_DATETIME] = '9000-01-01')
        ORDER BY DFL.START_DATETIME
    ''')

    sfl = min(list(filter(lambda dfl: re.search("s.{0,1}f.{0,1}l", str(dfl[2]), re.IGNORECASE) or dfl[3] != 'PRODUCING', list(cursor.fetchall()))), key= lambda dfl: dfl[1])
    print(sfl)
    PR = convert_FL_Pwf(all_well_test_data, sfl, mid_perf)



    # highest_FL = min(fluid_level_data, key = lambda dfl: dfl[1] if dfl[1] else dfl[3] if dfl[3] else math.inf)
    # highest_FL_2 = min(filter(lambda dfl: dfl[0] != highest_FL[0], fluid_level_data), key = lambda dfl: dfl[1] if dfl[1] else dfl[3] if dfl[3] else math.inf) or highest_FL
    # highest_FL_3 = min(filter(lambda dfl: dfl[0] != highest_FL[0] and dfl[0] != highest_FL_2[0], fluid_level_data), key = lambda dfl: dfl[1] if dfl[1] else dfl[3] if dfl[3] else math.inf) or highest_FL


    # PR = sum([convert_FL_Pwf(all_well_test_data, highest_FL, mid_perf), convert_FL_Pwf(all_well_test_data, highest_FL_2, mid_perf), convert_FL_Pwf(all_well_test_data, highest_FL_3, mid_perf)]) / 3





    y_train = np.array([])
    x_train = np.array([])
    for i in range(len(well_test_data)):
        closest_DFL = min(filter(lambda dfl: dfl[1], fluid_level_data), key = lambda dfl: abs(well_test_data[i][0] - dfl[0]))

        pwf = convert_FL_Pwf(all_well_test_data, closest_DFL, mid_perf)


        y_train = np.append(y_train, float(pwf))
        x_train = np.append(x_train, float(well_test_data[i][3]))

    plt.scatter(x_train,y_train)

    w_final, b_final, J_hist, p_hist = gradient_descent(x_train ,y_train, -1, PR, 0.00000005, 
                                                            10000, compute_cost, compute_gradient)


    print(f"(w,b) found by gradient descent: ({w_final:8.4f},{b_final:8.4f},{b_final/-w_final:8.4f})")
                                                         
    plt.plot([0, b_final/-w_final], [b_final, 0])
    plt.text(50, 250, f'Reservoir Pressure: {int(PR)}\nAOF: {int(b_final/-w_final)}')
    plt.show()
                                                        



def convert_FL_Pwf(all_well_test_data, closest_DFL, mid_perf):
    wc = min(all_well_test_data, key = lambda test: abs(closest_DFL[0] - test[0]))[4]

    cv_gamma = float(wc) * 0.45 + float(1 - float(wc)) * 0.35

    pwf = (float(mid_perf) - float(closest_DFL[1] if closest_DFL[1] else closest_DFL[3])) * cv_gamma

    return pwf
                




def compute_cost(x, y, w, b):
   
    m = x.shape[0] 
    cost = 0
    
    for i in range(m):
        f_wb = w * x[i] + b
        cost = cost + (f_wb - y[i])**2
    total_cost = 1 / (2 * m) * cost

    return total_cost


def compute_gradient(x, y, w, b): 
    """
    Computes the gradient for linear regression 
    Args:
      x (ndarray (m,)): Data, m examples 
      y (ndarray (m,)): target values
      w,b (scalar)    : model parameters  
    Returns
      dj_dw (scalar): The gradient of the cost w.r.t. the parameters w
      dj_db (scalar): The gradient of the cost w.r.t. the parameter b     
     """
    
    # Number of training examples
    m = x.shape[0]    
    dj_dw = 0
    dj_db = 0
    
    for i in range(m):  
        f_wb = w * x[i] + b 


        dj_dw_i = (f_wb - y[i]) * x[i] 
        dj_db_i = f_wb - y[i] 
        dj_db += dj_db_i
        dj_dw += dj_dw_i 
    dj_dw = dj_dw / m 
    dj_db = dj_db / m 
        
    return dj_dw, dj_db


def gradient_descent(x, y, w_in, b_in, alpha, num_iters, cost_function, gradient_function): 
    """
    Performs gradient descent to fit w,b. Updates w,b by taking 
    num_iters gradient steps with learning rate alpha
    
    Args:
      x (ndarray (m,))  : Data, m examples 
      y (ndarray (m,))  : target values
      w_in,b_in (scalar): initial values of model parameters  
      alpha (float):     Learning rate
      num_iters (int):   number of iterations to run gradient descent
      cost_function:     function to call to produce cost
      gradient_function: function to call to produce gradient
      
    Returns:
      w (scalar): Updated value of parameter after running gradient descent
      b (scalar): Updated value of parameter after running gradient descent
      J_history (List): History of cost values
      p_history (list): History of parameters [w,b] 
      """
    
    w = copy.deepcopy(w_in) # avoid modifying global w_in
    # An array to store cost J and w's at each iteration primarily for graphing later
    J_history = []
    p_history = []
    b = b_in
    w = w_in
    
    for i in range(num_iters):
        # Calculate the gradient and update the parameters using gradient_function
        dj_dw, dj_db = gradient_function(x, y, w , b)     

        # Update Parameters using equation (3) above
        # b = b - alpha * dj_db                            
        w = w - alpha * dj_dw                            

        # Save cost J at each iteration
        if i<100000:      # prevent resource exhaustion 
            J_history.append( cost_function(x, y, w , b))
            p_history.append([w,b])
        # Print cost every at intervals 10 times or as many iterations if < 10
        if i% math.ceil(num_iters/10) == 0:
            print(f"Iteration {i:4}: Cost {J_history[-1]:8.4f} ",
                  f"dj_dw: {dj_dw: 8.4f}, dj_db: {dj_db: 8.4f}  ",
                  f"w: {w: 8.4f}, b:{b: 8.4f}")
 
    return w, b, J_history, p_history #return w and J,w history for graphing

# while True:
#     main()

# Create object
root = Tk()
  
# Adjust size
root.geometry( "200x200" )
  
# Change the label text
def getIPR():
    main(clicked.get())

  
# Dropdown menu options


conn = pyodbc.connect('Driver={SQL Server};'
        'Server=QPCAVODB2K12;'
        'Database=AVOCET_PRODUCTION;'
        'Trusted_Connection=yes;'
        'UID=sa;'
        'PWD=AAAAAA;')

cursor = conn.cursor()
cursor.execute(f'''
    SELECT ITEM_NAME FROM VI_COMPLETION_en_US WHERE TYPE = 'PRODUCTION' ORDER BY ITEM_NAME
    ''')


options = map(lambda x: x[0], list(cursor.fetchall()))
  
# datatype of menu text
clicked = StringVar()
  

  
# Create Dropdown menu
drop = OptionMenu( root , clicked , *options )
drop.pack()
  
# Create button, it will change label text
button = Button( root , text = "Calculate IPR" , command = getIPR ).pack()
  
# Create Label

  
# Execute tkinter
root.mainloop()

main()