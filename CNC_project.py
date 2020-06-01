import math
import matplotlib.pyplot as plt
import numpy as np
from xlwt import Workbook
wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

sheet1.write(0,0,'Job')
sheet1.write(0,1,'Operation')
sheet1.write(0,2,'Maximum RPM')
sheet1.write(0,3,'Cutting Speed (m/min)')
sheet1.write(0,4,'Feed/rev (mm/rev)')
sheet1.write(0,5,'Feed/min (mm/min)')
sheet1.write(0,6,'Diameter @ max RPM (mm)')
sheet1.write(0,7,'Cutting time/stroke (min)')
sheet1.write(0,8,'Metal removal rate (cm^3/min)')
sheet1.write(0,9,'Probable cycle time (min)')
sheet1.write(0,10,'No. of jobs')
sheet1.write(0,11,'Rework qty.')
sheet1.write(0,12,'Rejected qty.')
sheet1.write(0,13,'Despatch qty.')
sheet1.write(0,14,'Loading/Unloading time (sec)')
sheet1.write(0,15,'Setting & other Miscellaneous time (min)')
sheet1.write(0,16,'Active machining time (hrs.)')
sheet1.write(0,17,'Total machining time (hrs.)')
sheet1.write(0,18,'Electricity consumption cost (Rs.)')
sheet1.write(0,19,'Manpower cost (Rs.)')
sheet1.write(0,20,'No. of inserts required')
sheet1.write(0,21,'Insert cost (Rs.)')
sheet1.write(0,22,'Total insert cost (Rs.)')
sheet1.write(0,23,'Miscellaneous cost (Rs.)')
sheet1.write(0,24,'Total machining cost (Rs.)')
sheet1.write(0,25,'Cost/job (Rs.)')
sheet1.write(0,26,'Offer Rate/job (Rs.)')
sheet1.write(0,27,'Total S.P. (Rs.)')
sheet1.write(0,28,'Profit (Rs.)')
sheet1.write(0,29,'Profit %')
wb.save('CNC Job Quotation.xls')



no_oprn = int(input('Enter no. of Turning operations to perform= '))
no_opn = no_oprn
no_oprns = range(1, (no_oprn+1))

act_totMC_cost = []
act_Tmch_time = []
act_totIns_cost = []




for pieces in no_oprns:
    while no_oprn == no_opn:
        print(f'Operation no. {pieces}')
        auto_cal = input('Do you want a (A)utomated or (M)anual entry? ')
        opr_name = input("Enter Job name and operation description: ")
        sheet1.write(3*pieces,0,opr_name)

        if auto_cal.upper() == 'A':

            oprn = input('Do you want to include both Roughing & Finishing operations? (Y/N): ')
            sheet1.write(3*pieces,1,'Rough')
            sheet1.write(3*pieces+1,1,'Finish')
            
            if oprn.upper() == 'Y':               
                max_rpm = int(input("Enter maximum RPM for Roughing= "))  
                vc = int(input("Enter Cutting speed for Roughing (m/min)= "))
                feed_rev = float(input("Enter feed rate/revolution for Roughing (mm/rev)= "))
                dia = float(input("Enter mean diameter (mm)= "))          
                max_rpmF = int(input("Enter maximum RPM for Finishing= "))
                vcF = int(input("Enter Cutting speed for Finishing (m/min)= "))
                feed_revF = float(input("Enter feed rate/revolution for Finishing (mm/rev)= "))
                rpm = float(vc/dia*1000/3.14)
                rpmF = float(vcF/dia*1000/3.14)
                #print(f'RPM for Roughing= {round(rpm)}')
                #print(f'RPM for Finishing= {round(rpmF)}')
                sheet1.write(3*pieces, 2, max_rpm)
                sheet1.write(3*pieces, 3, vc)
                sheet1.write(3*pieces, 4, feed_rev)            
                sheet1.write(3*pieces+1, 2, max_rpmF)
                sheet1.write(3*pieces+1, 3, vcF)
                sheet1.write(3*pieces+1, 4, feed_revF)
                
                
                if rpm>max_rpm:
                    feed_min = int(feed_rev * max_rpm)
                    print(f"Feed/min for Roughing= {feed_min} mm/min")
                else:
                    feed_min = int(feed_rev * rpm)
                    print(f"Feed/min for Roughing= {feed_min} mm/min")

                if rpmF>max_rpmF:
                    feed_minF = int(feed_revF * max_rpmF)
                    print(f"Feed/min for Finishing= {feed_minF} mm/min")
                else:
                    feed_minF = int(feed_revF * rpmF)
                    print(f"Feed/min for Finishing= {feed_minF} mm/min")

                sheet1.write(3*pieces, 5, feed_min)
                sheet1.write(3*pieces+1, 5, feed_minF)
                

                dia_max_rpm = float(vc/max_rpm*1000/3.14)
                print(f'Diameter at which maximum RPM is achieved in Roughing= {round(dia_max_rpm)} mm')

                dia_max_rpmF = float(vcF/max_rpmF*1000/3.14)
                print(f'Diameter at which maximum RPM is achieved in Finishing= {round(dia_max_rpmF)} mm')
                sheet1.write(3*pieces, 6, max_rpm)
                sheet1.write(3*pieces+1, 6, max_rpmF)

                den_steel = 7.8          #g/cm^3

                ask_den = int(input('''Enter the serial no. of material for density:
                
                1. steel = 7.8 g/cm^3
                M. (Manual entry) 

                '''))

                if ask_den == 1:
                    mat_den = den_steel

                elif ask_den.upper() == 'M':
                    mat_den = float(input("Enter work material density manually (g/cm^3)= "))
                
                mat_len = float(input("Enter work material length (mm)= "))
                
                CT_stk = float(mat_len/feed_min)
                print(f'Cutting time/stroke for Roughing= {CT_stk} min')

                CT_stkF = float(mat_len/feed_minF)
                print(f'Cutting time/stroke for Finishing= {CT_stkF} min')
                sheet1.write(3*pieces, 7, CT_stk)
                sheet1.write(3*pieces+1, 7, CT_stkF)
                
                cut_depth = float(input("Enter cutting depth for Roughing (diameter value in mm)= "))
                fin_stk = float(input("Enter finish stock (radius value in mm)= "))

                Q = float(vc*cut_depth*feed_rev)
                print(f'Metal removal rate in Roughing= {Q} cm^3/min')
                Q_cut = float(Q*CT_stk)
                print(f'Metal removed in a single cut in Roughing= {Q_cut} cm^3')

                

                QF = float(vcF*2*fin_stk*feed_revF)
                print(f'Metal removal rate in Finishing= {QF} cm^3/min')
                Q_cutF = float(QF*CT_stkF)
                print(f'Metal removed in a single cut in Finishing= {Q_cutF} cm^3')

                sheet1.write(3*pieces, 8, Q_cut)
                sheet1.write(3*pieces+1, 8, Q_cutF)

                job_type = input('''
            Enter job type:
            Uniform jobs (U)
            Non-uniform jobs (N)''')

                if job_type.upper() == "U":
                    print(".....UNIFORM JOB.....")

                    init_dia = float(input("Enter initial diameter of job(mm)= "))
                    final_dia = float(input("Enter final diameter of job(mm)= "))
                    bore_dia = float(input("Enter bore diameter of job(mm)= "))

                    init_vol = float((3.14*(init_dia**2-bore_dia**2)*mat_len)/4000)
                    print(f'Initial volume= {round(init_vol)} cm^3')   #
                    
                    no_cutsR = float(init_dia/cut_depth - final_dia/cut_depth-(2*fin_stk)/cut_depth) 
                    no_cuts = float(init_dia/cut_depth - final_dia/cut_depth)
                    print(f'Total no. of Roughing cuts= {math.ceil(no_cutsR)}')
                    print(f'Total no. of Finishing cuts= {math.ceil(no_cuts)}')

                    final_volR = float((3.14*((final_dia + (2*fin_stk))**2-bore_dia**2)*mat_len)/4000)
                    final_vol = float((3.14*(final_dia**2-bore_dia**2)*mat_len)/4000)
                    print(f'Final volume after Roughing= {round(final_volR)} cm^3')
                    print(f'Final volume after Finishing= {round(final_vol)} cm^3')

                    #fin_stk = float(input("Enter finish stock (radius value in mm)= "))

                    


                elif job_type.upper() == "N":
                    print(".....NON-UNIFORM JOB.....")

                    no_sec = int(input("Enter no. of sections in the job= "))
                    no_dia = int(no_sec+2)
                    sec = range(1, no_dia)      
                    lst1 = []
                    for dias in sec:
                        diam = float(input(f"Enter diameter {dias} (mm)= "))
                        lst1.append(diam)
                    while len(lst1)<=10:
                        lst1.append(0.0)
                    #print(lst1)
                    
                    
                    
                    
                    sec_len = range(1, (no_sec+1))
                    lst2 = []    
                    for lens in sec_len:
                        length = float(input(f"Enter length of {lens} section (mm)= "))
                        lst2.append(length)
                    while len(lst2)<=10:
                        lst2.append(0.0)
                    #print(lst2)

                    #fin_stk = float(input("Enter finish stock (radius value in mm)= "))

                    a, b, c, d, e, f, g, h, i, j, k =  lst1
                    p, q, r, s, t, u, v, w, x, y, z = lst2

                    
                    

                    Fvol1 = ((3.14*p*((a**2)+(b**2)+(a*b)))/12)
                    Fvol2 = ((3.14*q*((b**2)+(c**2)+(b*c)))/12)
                    Fvol3 = ((3.14*r*((c**2)+(d**2)+(c*d)))/12)
                    Fvol4 = ((3.14*s*((d**2)+(e**2)+(d*e)))/12)
                    Fvol5 = ((3.14*t*((e**2)+(f**2)+(e*f)))/12)
                    Fvol6 = ((3.14*u*((f**2)+(g**2)+(f*g)))/12)
                    Fvol7 = ((3.14*v*((g**2)+(h**2)+(g*h)))/12)
                    Fvol8 = ((3.14*w*((h**2)+(i**2)+(h*i)))/12)
                    Fvol9 = ((3.14*x*((i**2)+(j**2)+(i*j)))/12)
                    Fvol10 = ((3.14*y*((j**2)+(k**2)+(j*k)))/12)

                    final_Fvol = float((Fvol1 + Fvol2 + Fvol3 + Fvol4 + Fvol5 + Fvol6 + Fvol7 + Fvol8 + Fvol9 + Fvol10)/1000)


                    Rvol1 = (Fvol1+((3.14*p*(2*fin_stk)*((2*fin_stk)+b+a))/4))
                    Rvol2 = (Fvol2+((3.14*q*(2*fin_stk)*((2*fin_stk)+c+b))/4))
                    Rvol3 = (Fvol3+((3.14*r*(2*fin_stk)*((2*fin_stk)+d+c))/4))
                    Rvol4 = (Fvol4+((3.14*s*(2*fin_stk)*((2*fin_stk)+e+d))/4))
                    Rvol5 = (Fvol5+((3.14*t*(2*fin_stk)*((2*fin_stk)+f+e))/4))
                    Rvol6 = (Fvol6+((3.14*u*(2*fin_stk)*((2*fin_stk)+g+f))/4))
                    Rvol7 = (Fvol7+((3.14*v*(2*fin_stk)*((2*fin_stk)+h+g))/4))
                    Rvol8 = (Fvol8+((3.14*w*(2*fin_stk)*((2*fin_stk)+i+h))/4))
                    Rvol9 = (Fvol9+((3.14*x*(2*fin_stk)*((2*fin_stk)+j+i))/4))
                    Rvol10 = (Fvol10+((3.14*y*(2*fin_stk)*((2*fin_stk)+k+j))/4))

                    final_Rvol = float((Rvol1 + Rvol2 + Rvol3 + Rvol4 + Rvol5 + Rvol6 + Rvol7 + Rvol8 + Rvol9 + Rvol10)/1000)




                    init_diam = float(input("Enter initial diameter of work material (mm)= "))
                    bore_diam = float(input("Enter initial bore diameter of work material (mm)= "))
                    init_vol = float((3.14*(init_diam**2-bore_diam**2)*mat_len)/4000)

                    print(f'Total initial volume = {init_vol} cm^3')
                    print(f'Rough volume of sections (cm^3)= {[Rvol1, Rvol2, Rvol3, Rvol4, Rvol5, Rvol6, Rvol7, Rvol8, Rvol9, Rvol10]}')
                    print(f'Finish volume of sections (cm^3)= {[Fvol1, Fvol2, Fvol3, Fvol4, Fvol5, Fvol6, Fvol7, Fvol8, Fvol9, Fvol10]}')
                    print(f'Volume after Roughing= {final_Rvol} cm^3')
                    print(f'Finished Volume= {final_Fvol} cm^3')

                    
                else:
                    print("!! Invalid Entry !!")



                Rmat_rmv = init_vol-final_Rvol
                Fmat_rmv = final_Rvol-final_Fvol
                mat_rmv = Rmat_rmv + Fmat_rmv
                #print(f'Material removed in Roughing= {Rmat_rmv} cm^3')
                #print(f'Material removed in Finishing= {Fmat_rmv} cm^3')
                print(f'Total material removed= {mat_rmv} cm^3')
                print(f'Percentage of material removed= {mat_rmv / init_vol * 100}%')

                auto_inputCT = input("Do you want (A)utomated or (M)anual cycle time entry: ")
                if auto_inputCT.upper() == 'A':
                    if job_type.upper() == "N":
                        probCT_R = float(Rmat_rmv/Q)
                        probCT_F = float(Fmat_rmv/QF)
                        probCT = float(probCT_R + probCT_F)
                    else:
                        probCT_R = float(CT_stk * no_cutsR)
                        probCT_F = float(CT_stk * no_cuts)
                        probCT = float(probCT_R + probCT_F)
                elif auto_inputCT.upper() == 'M':
                    probCT_R = float(input('Enter Roughing cycle time (min.)= '))
                    probCT_F = float(input('Enter Roughing cycle time (min.)= '))
                    probCT = float(probCT_R + probCT_F)
                else:
                    print("...Invalid Input...")
                    
                print(f'Probable cycle time= {probCT} min')

                sheet1.write(3*pieces, 9, probCT)

                already_rej_qty = int(input("Total number of already rejected jobs (if any)= "))
                jobs = int(input("Total number of jobs= "))
                jobs = int(jobs-already_rej_qty)
                rw_qty = int(input("Total number of reworked jobs= "))
                rej_qty = int(input("Total number of rejected jobs= "))
                des_qty = int(jobs-rej_qty)
                LU_time = float(input("Enter Loading/Unloading time (seconds)=  "))
                set_time = float(input("Enter Setting & Miscellanious time (min) "))
                Amch_time = float(probCT*(jobs+rw_qty)/60)
                Tmch_time = float((LU_time*jobs/3600)+Amch_time+(set_time/60))
                act_Tmch_time.append(Tmch_time)

                Amch_time_R = float(probCT_R*(jobs+rw_qty)/60)
                Amch_time_F = float(probCT_F*(jobs+rw_qty)/60)

                #print(f'Active Machining time= {Amch_time}')
                print(f'Total Machining time= {Tmch_time} hrs.')

                sheet1.write(3*pieces, 10, jobs)
                sheet1.write(3*pieces, 11, rw_qty)
                sheet1.write(3*pieces, 12, rej_qty)
                sheet1.write(3*pieces, 13, des_qty)
                sheet1.write(3*pieces, 14, LU_time)
                sheet1.write(3*pieces, 15, set_time)
                sheet1.write(3*pieces, 16, Amch_time)
                sheet1.write(3*pieces, 17, Tmch_time)

                mchKW = 7.5 #kW
                elec1U = 8  #Rs./unit

                totKW = float(Tmch_time * mchKW)
                kw_cost = float(totKW * elec1U)

                print(f'Electricity consumption cost= Rs.{kw_cost}')
                sheet1.write(3*pieces, 18, kw_cost)

                opt_cost = float(input("Enter operator cost/hr.= Rs."))
                manP_cost = float(opt_cost * Tmch_time)
                print(f'Manpower cost= Rs.{manP_cost}')
                sheet1.write(3*pieces, 19, manP_cost)

                RE = float(input("Enter Roughing insert corner radius (mm)= "))
                ins_corner = int(input('Enter no. of corners in the Roughing insert= '))
                TSR = float(feed_rev**2/RE*1000/8) 
                #sheet1.write(3*pieces, , RE)  

                C = float(input('C= '))
                X = float(input('X= '))
                Y = float(input('Y= '))
                N = float(input('N= '))
                corner_life = float((C/(((cut_depth/1000)**X)*((feed_min/1000)**Y)*vc))**(1/N))
                ins_life = float(corner_life * ins_corner)
                #sheet1.write(3*pieces, , ins_life)

                ins_req = int(math.ceil(Amch_time_R / ins_life))
                ins_cost = float(input('Enter price of 1 Roughing insert= Rs.'))
                totIns_costR = float(ins_req * ins_cost)
                print(f'Total surface roughness after Roughing= {TSR}')
                print(f'Roughing insert life is {ins_life} min')
                print(f'Number of Roughing inserts required= {ins_req}')
                print(f'Total Roughing insert cost= Rs.{totIns_costR}')
                sheet1.write(3*pieces, 20, ins_req)
                sheet1.write(3*pieces, 21, totIns_costR)

                RE_F = float(input("Enter Finishing insert corner radius (mm)= "))
                ins_cornerF = int(input('Enter no. of corners in the Finishing insert= '))
                TSR_F = float(feed_revF**2/RE_F*1000/8)
                #sheet1.write(3*pieces+1, , RE_F)   

                C_F = float(input('C= '))
                X_F = float(input('X= '))
                Y_F = float(input('Y= '))
                N_F = float(input('N= '))
                corner_lifeF = float((C_F/((((2*fin_stk)/1000)**X_F)*((feed_minF/1000)**Y_F)*vcF))**(1/N_F))
                ins_lifeF = float(corner_lifeF * ins_cornerF)
                #sheet1.write(3*pieces+1, , ins_lifeF)

                ins_reqF = int(math.ceil(Amch_time_F / ins_lifeF))
                ins_costF = float(input('Enter price of 1 Finishing insert= Rs.'))
                totIns_costF = float(ins_reqF * ins_costF)
                print(f'Total surface roughness after Finishing= {TSR_F}')
                print(f'Finishing insert life is {ins_lifeF} min')
                print(f'Number of Finishing inserts required= {ins_reqF}')
                print(f'Total Finishing insert cost= Rs.{totIns_costF}')
                sheet1.write(3*pieces+1, 20, ins_reqF)
                sheet1.write(3*pieces+1, 21, totIns_costF)

                totIns_cost = totIns_costR + totIns_costF
                print(f'Total insert cost= Rs.{totIns_cost}')
                act_totIns_cost.append(totIns_cost)
                sheet1.write(3*pieces, 22, totIns_cost)


                totMC_cost = float(kw_cost + manP_cost + totIns_cost)
                print(f'Total production expense/setting= Rs.{totMC_cost}')
                act_totMC_cost.append(totMC_cost)

                break            
            










            elif oprn.upper() == 'N':
                max_rpm = int(input("Enter maximum RPM= "))  

                arbit_max_rpm = np.array(range(1, 5000))

                vc = int(input("Enter Cutting speed (m/min)= "))

                arbit_vc = np.array(range(1, 1000))

                feed_rev = float(input("Enter feed rate/revolution (mm/rev)= "))
                arbit_feed_rev1 = np.array(range(1, 10))
                arbit_feed_rev = arbit_feed_rev1 * 0.1
                arbit_feed_rev2 = np.array(range(1, 1000))

                dia = float(input("Enter mean diameter (mm)= "))            
                rpm = float(vc/dia*1000/3.14)  
                arbit_rpm2 = arbit_vc/dia*1000/3.14

                #print(f'RPM for Roughing= {round(rpm)}')            
                sheet1.write(3*pieces, 2, max_rpm)
                sheet1.write(3*pieces, 3, vc)
                sheet1.write(3*pieces, 4, feed_rev)            
                
                
                
                if rpm>max_rpm:
                    feed_min = int(feed_rev * max_rpm)
                    print(f"Feed/min for= {feed_min} mm/min")
                    arbit_feed_min1 = arbit_feed_rev * max_rpm
                    arbit_feed_min2_1 = arbit_feed_rev2 * max_rpm
                else:
                    feed_min = int(feed_rev * rpm)
                    print(f"Feed/min= {feed_min} mm/min")
                    arbit_feed_min1 = arbit_feed_rev * rpm
                    arbit_feed_min2 = arbit_feed_rev2 * arbit_rpm2
                    
                


                sheet1.write(3*pieces, 5, feed_min)
                

                dia_max_rpm = float(vc/max_rpm*1000/3.14)
                print(f'Diameter at which maximum RPM is achieved = {round(dia_max_rpm)} mm')

                sheet1.write(3*pieces, 6, feed_min)
                
                den_steel = 7.8          #g/cm^3

                ask_den = int(input('''Enter the serial no. of material for density:
                
                1. steel = 7.8 g/cm^3
                M. (Manual entry) 

                '''))

                if ask_den == 1:
                    mat_den = den_steel

                elif ask_den.upper() == 'M':
                    mat_den = float(input("Enter work material density manually (g/cm^3)= "))
                
                mat_len = float(input("Enter work material length (mm)= "))
                
                CT_stk = float(mat_len/feed_min)
                print(f'Cutting time/stroke= {CT_stk} min')
                arbit_CT_stk1 = mat_len/arbit_feed_min1

                if rpm>max_rpm:
                    arbit_CT_stk2 = mat_len/arbit_feed_min2_1
                else:
                    arbit_CT_stk2 = mat_len/arbit_feed_min2

                sheet1.write(3*pieces, 7, CT_stk)
                
                
                cut_depth = float(input("Enter cutting depth (diameter value in mm)= "))
                arbit_cut_depth = np.array(range(1, 10))
                

                Q = float(vc*cut_depth*feed_rev)
                print(f'Metal removal rate= {Q} cm^3/min')
                arbit_Q1 = vc*cut_depth*arbit_feed_rev
                arbit_Q2 = arbit_vc*cut_depth*feed_rev
                Q_cut = float(Q*CT_stk)
                print(f'Metal removed in a single cut= {Q_cut} cm^3')
                arbit_Q_cut1 = arbit_Q1*arbit_CT_stk1
                arbit_Q_cut2 = arbit_Q2*arbit_CT_stk2

                sheet1.write(3*pieces, 8, Q_cut)

                

                job_type = input('''
            Enter job type:
            Uniform jobs (U)
            Non-uniform jobs (N)''')

                if job_type.upper() == "U":
                    print(".....UNIFORM JOB.....")

                    init_dia = float(input("Enter initial diameter of job(mm)= "))
                    final_dia = float(input("Enter final diameter of job(mm)= "))
                    bore_dia = float(input("Enter bore diameter of job(mm)= "))

                    init_vol = float((3.14*(init_dia**2-bore_dia**2)*mat_len)/4000)
                    print(f'Initial volume= {round(init_vol)} cm^3')

                    final_vol = float((3.14*(final_dia**2-bore_dia**2)*mat_len)/4000)
                    print(f'Final volume= {round(final_vol)} cm^3')

                    no_cuts = int(init_dia/cut_depth - final_dia/cut_depth)
                    print(f'Total no. of cuts= {no_cuts}')
                    arbit_no_cuts = init_dia/cut_depth - final_dia/arbit_cut_depth

                    #fin_stk = float(input("Enter finish stock (radius value in mm)= "))

                    


                elif job_type.upper() == "N":
                    print(".....NON-UNIFORM JOB.....")

                    no_sec = int(input("Enter no. of sections in the job= "))
                    no_dia = int(no_sec+2)
                    sec = range(1, no_dia)      
                    lst1 = []
                    for dias in sec:
                        diam = float(input(f"Enter diameter {dias} (mm)= "))
                        lst1.append(diam)
                    while len(lst1)<=10:
                        lst1.append(0.0)
                    #print(lst1)
                    
                    
                    
                    
                    sec_len = range(1, (no_sec+1))
                    lst2 = []    
                    for lens in sec_len:
                        length = float(input(f"Enter length of {lens} section (mm)= "))
                        lst2.append(length)
                    while len(lst2)<=10:
                        lst2.append(0.0)
                    #print(lst2)

                    #fin_stk = float(input("Enter finish stock (radius value in mm)= "))

                    a, b, c, d, e, f, g, h, i, j, k =  lst1
                    p, q, r, s, t, u, v, w, x, y, z = lst2

                    
                    

                    Fvol1 = ((3.14*p*((a**2)+(b**2)+(a*b)))/12)
                    Fvol2 = ((3.14*q*((b**2)+(c**2)+(b*c)))/12)
                    Fvol3 = ((3.14*r*((c**2)+(d**2)+(c*d)))/12)
                    Fvol4 = ((3.14*s*((d**2)+(e**2)+(d*e)))/12)
                    Fvol5 = ((3.14*t*((e**2)+(f**2)+(e*f)))/12)
                    Fvol6 = ((3.14*u*((f**2)+(g**2)+(f*g)))/12)
                    Fvol7 = ((3.14*v*((g**2)+(h**2)+(g*h)))/12)
                    Fvol8 = ((3.14*w*((h**2)+(i**2)+(h*i)))/12)
                    Fvol9 = ((3.14*x*((i**2)+(j**2)+(i*j)))/12)
                    Fvol10 = ((3.14*y*((j**2)+(k**2)+(j*k)))/12)

                    final_Fvol = float((Fvol1 + Fvol2 + Fvol3 + Fvol4 + Fvol5 + Fvol6 + Fvol7 + Fvol8 + Fvol9 + Fvol10)/1000)


                    




                    init_diam = float(input("Enter initial diameter of work material (mm)= "))
                    bore_diam = float(input("Enter initial bore diameter of work material (mm)= "))
                    init_vol = float((3.14*(init_diam**2-bore_diam**2)*mat_len)/4000)

                    print(f'Total initial volume = {init_vol} cm^3')
                    print(f'Volume of sections (cm^3)= {[Fvol1, Fvol2, Fvol3, Fvol4, Fvol5, Fvol6, Fvol7, Fvol8, Fvol9, Fvol10]}')                
                    print(f'Final Volume= {final_Fvol} cm^3')

                    
                else:
                    print("!! Invalid Entry !!")


                
                mat_rmv = init_vol-final_Fvol           
                print(f'Total material removed= {mat_rmv} cm^3')
                print(f'Percentage of material removed= {mat_rmv / init_vol * 100}%')

                auto_inputCT = input("Do you want (A)utomated or (M)anual cycle time entry: ")
                if auto_inputCT.upper() == 'A':
                    if job_type.upper() == "N":                    
                        probCT = float(mat_rmv/Q)
                        arbit_probCT1 = mat_rmv/arbit_Q1
                        arbit_probCT2 = mat_rmv/arbit_Q2
                    else:
                        probCT = float(CT_stk * no_cuts)
                        arbit_probCT1 = arbit_CT_stk1 * no_cuts
                        arbit_probCT2 = arbit_CT_stk2 * no_cuts
                elif auto_inputCT.upper() == 'M':
                    probCT= float(input('Enter cycle time (min.)= '))                
                else:
                    print("...Invalid Input...")

            
                    
                print(f'Probable cycle time= {probCT} min')

                sheet1.write(3*pieces, 9, probCT)

                already_rej_qty = int(input("Total number of already rejected jobs (if any)= "))
                jobs = int(input("Total number of jobs= "))
                jobs = int(jobs-already_rej_qty)
                arbit_jobs = np.array(range(1, 100))
                rw_qty = int(input("Total number of reworked jobs= "))
                arbit_rw_qty = np.array(range(1, 100))
                rej_qty = int(input("Total number of rejected jobs= "))
                arbit_rej_qty = np.array(range(1, 100))
                des_qty = int(jobs-rej_qty)
                LU_time = float(input("Enter Loading/Unloading time (seconds)=  "))
                set_time = float(input("Enter Setting & Miscellanious time (min) "))
                Amch_time = float(probCT*(jobs+rw_qty)/60)
                arbit_Amch_time1 = arbit_probCT1*(jobs+rw_qty)/60
                arbit_Amch_time2 = arbit_probCT2*(jobs+rw_qty)/60
                arbit_Amch_time3 = probCT*(arbit_jobs+rw_qty)/60
                arbit_Amch_time4 = probCT*(jobs+arbit_rw_qty)/60
                Tmch_time = float((LU_time*jobs/3600)+Amch_time+(set_time/60))
                arbit_Tmch_time1 = (LU_time*jobs/3600)+arbit_Amch_time1+(set_time/60)
                arbit_Tmch_time2 = (LU_time*jobs/3600)+arbit_Amch_time2+(set_time/60)
                arbit_Tmch_time3 = (LU_time*jobs/3600)+arbit_Amch_time3+(set_time/60)
                arbit_Tmch_time4 = (LU_time*jobs/3600)+arbit_Amch_time4+(set_time/60)
                act_Tmch_time.append(Tmch_time)

                #print(f'Active Machining time= {Amch_time}')
                print(f'Total Machining time= {Tmch_time} hrs.')

                sheet1.write(3*pieces, 10, jobs)
                sheet1.write(3*pieces, 11, rw_qty)
                sheet1.write(3*pieces, 12, rej_qty)
                sheet1.write(3*pieces, 13, des_qty)
                sheet1.write(3*pieces, 14, LU_time)
                sheet1.write(3*pieces, 15, set_time)
                sheet1.write(3*pieces, 16, Amch_time)
                sheet1.write(3*pieces, 17, Tmch_time)

                mchKW = 7.5 #kW
                elec1U = 8  #Rs./unit

                totKW = float(Tmch_time * mchKW)
                arbit_totKW1 = arbit_Tmch_time1 * mchKW
                arbit_totKW2 = arbit_Tmch_time2 * mchKW
                arbit_totKW3 = arbit_Tmch_time3 * mchKW
                arbit_totKW4 = arbit_Tmch_time4 * mchKW
                kw_cost = float(totKW * elec1U)
                arbit_kw_cost1 = arbit_totKW1 * elec1U
                arbit_kw_cost2 = arbit_totKW2 * elec1U
                arbit_kw_cost3 = arbit_totKW3 * elec1U
                arbit_kw_cost4 = arbit_totKW4 * elec1U

                print(f'Electricity consumption cost= Rs.{kw_cost}')
                sheet1.write(3*pieces, 18, kw_cost)

                opt_cost = float(input("Enter operator cost/hr.= Rs."))
                manP_cost = float(opt_cost * Tmch_time)
                arbit_manP_cost1 = opt_cost * arbit_Tmch_time1
                arbit_manP_cost2 = opt_cost * arbit_Tmch_time2
                arbit_manP_cost3 = opt_cost * arbit_Tmch_time3
                arbit_manP_cost4 = opt_cost * arbit_Tmch_time4
                print(f'Manpower cost cost= Rs.{manP_cost}')
                sheet1.write(3*pieces, 19, manP_cost)

                RE = float(input("Enter insert corner radius (mm)= "))
                ins_corner = int(input('Enter no. of corners in the insert= '))
                TSR = float(feed_rev**2/RE*1000/8) 
                
                C = float(input('C= '))
                X = float(input('X= '))
                Y = float(input('Y= '))
                N = float(input('N= '))
                corner_life = float((C/(((cut_depth/1000)**X)*((feed_min/1000)**Y)*vc))**(1/N))
                arbit_corner_life1 = (C/(((cut_depth/1000)**X)*((arbit_feed_min1/1000)**Y)*vc))**(1/N)

                if rpm>max_rpm:
                    arbit_corner_life2 = (C/(((cut_depth/1000)**X)*((arbit_feed_min2_1/1000)**Y)*arbit_vc))**(1/N)
                else:
                    arbit_corner_life2 = (C/(((cut_depth/1000)**X)*((arbit_feed_min2/1000)**Y)*arbit_vc))**(1/N)


                
                ins_life = float(corner_life * ins_corner)
                arbit_ins_life1 = arbit_corner_life1 * ins_corner
                arbit_ins_life2 = arbit_corner_life2 * ins_corner
                
                ins_req = int(math.ceil(Amch_time / ins_life))
                arbit_ins_req1 = arbit_Amch_time1 / arbit_ins_life1
                arbit_ins_req2 = arbit_Amch_time2 / arbit_ins_life2
                arbit_ins_req3 = arbit_Amch_time3 / ins_life
                arbit_ins_req4 = arbit_Amch_time4 / ins_life
                ins_cost = float(input('Enter price of 1 Roughing insert= Rs.'))
                totIns_cost = float(ins_req * ins_cost)
                arbit_totIns_cost1 = arbit_ins_req1 * ins_cost
                arbit_totIns_cost2 = arbit_ins_req2 * ins_cost
                arbit_totIns_cost3 = arbit_ins_req3 * ins_cost
                arbit_totIns_cost4 = arbit_ins_req4 * ins_cost
                act_totIns_cost.append(totIns_cost)
                print(f'Total surface roughness= {TSR}')
                print(f'Roughing insert life is {ins_life} min')
                print(f'Number of inserts required= {ins_req}')
                print(f'Total insert cost= Rs.{totIns_cost}')
                sheet1.write(3*pieces, 20, ins_req)
                sheet1.write(3*pieces, 22, totIns_cost)    

                totMC_cost = float(kw_cost + manP_cost + totIns_cost)
                arbit_totMC_cost1 = arbit_kw_cost1 + arbit_manP_cost1 + arbit_totIns_cost1
                arbit_totMC_cost2 = arbit_kw_cost2 + arbit_manP_cost2 + arbit_totIns_cost2
                arbit_totMC_cost3 = arbit_kw_cost3 + arbit_manP_cost3 + arbit_totIns_cost3
                arbit_totMC_cost4 = arbit_kw_cost4 + arbit_manP_cost4 + arbit_totIns_cost4
                print(f'Total production expense/setting= Rs.{totMC_cost}')
                act_totMC_cost.append(totMC_cost)

                
                
                
               

                break









        elif auto_cal.upper() == 'M':            
        
            probCT_other= float(input('Enter cycle time (min.)= '))
            already_rej_qty_other = int(input("Total number of already rejected jobs (if any)= "))
            jobs_other = int(input("Total number of jobs= "))
            jobs_other = int(jobs_other-already_rej_qty_other)
            rw_qty_other = int(input("Total number of reworked jobs= "))
            rej_qty_other = int(input("Total number of rejected jobs= "))
            des_qty_other = int(jobs_other-rej_qty_other)
            LU_time_other = float(input("Enter Loading/Unloading time (seconds)=  "))
            set_time_other = float(input("Enter Setting & Miscellanious time (min) "))
            Amch_time_other = float(probCT_other*(jobs_other+rw_qty_other)/60)
            Tmch_time_other = float((LU_time_other*jobs_other/3600)+Amch_time_other+(set_time_other/60))
            act_Tmch_time.append(Tmch_time_other)

            sheet1.write(3*pieces, 9, probCT_other)
            sheet1.write(3*pieces, 10, jobs_other)
            sheet1.write(3*pieces, 11, rw_qty_other)
            sheet1.write(3*pieces, 12, rej_qty_other)
            sheet1.write(3*pieces, 13, des_qty_other)
            sheet1.write(3*pieces, 14, LU_time_other)
            sheet1.write(3*pieces, 15, set_time_other)
            sheet1.write(3*pieces, 16, Amch_time_other)
            sheet1.write(3*pieces, 17, Tmch_time_other)


            mchKW = 7.5 #kW
            elec1U = 8  #Rs./unit

            totKW_other = float(Tmch_time_other * mchKW)
            kw_cost_other = float(totKW_other * elec1U)

            print(f'Electricity consumption cost= Rs.{kw_cost_other}')
            sheet1.write(3*pieces, 18, kw_cost_other)

            opt_cost_other = float(input("Enter operator cost/hr.= Rs."))
            manP_cost_other = float(opt_cost_other * Tmch_time_other)
            print(f'Manpower cost cost= Rs.{manP_cost_other}')
            sheet1.write(3*pieces, 19, manP_cost_other)

            auto_input = input('Do you want a (A)utomated or (M)anual insert calculation? ')

            if auto_input.upper() == 'A':
                max_rpm_other = int(input("Enter maximum RPM= "))  
                vc_other = int(input("Enter Cutting speed (m/min)= "))
                feed_rev_other = float(input("Enter feed rate/revolution (mm/rev)= "))
                dia_other = float(input("Enter mean diameter (mm)= "))
                rpm_other= float(vc_other/dia_other*1000/3.14) 
                cut_depth_other = float(input("Enter cutting depth (diameter value in mm)= "))
                if rpm_other>max_rpm_other:
                    feed_min_other = int(feed_rev_other * max_rpm_other)
                    print(f"Feed/min for= {feed_min_other} mm/min")
                else:
                    feed_min_other = int(feed_rev_other * rpm_other)
                    print(f"Feed/min= {feed_min_other} mm/min")
                RE_other = float(input("Enter insert corner radius (mm)= "))
                ins_corner_other = int(input('Enter no. of corners in the insert= '))
                TSR_other = float(feed_rev_other**2/RE_other*1000/8) 
                
                C = float(input('C= '))
                X = float(input('X= '))
                Y = float(input('Y= '))
                N = float(input('N= '))
                corner_life_other = float((C/(((cut_depth_other/1000)**X)*((feed_min_other/1000)**Y)*vc_other))**(1/N))
                ins_life_other = float(corner_life_other * ins_corner_other)
                
                ins_req_other = int(math.ceil(Amch_time_other / ins_life_other))
                ins_cost_other = float(input('Enter price of 1 insert= Rs.'))
                totIns_cost_other = float(ins_req_other * ins_cost_other)
                act_totIns_cost.append(totIns_cost_other)
                print(f'Total surface roughness= {TSR_other}')
                print(f'Roughing insert life is {ins_life_other} min')
                print(f'Number of inserts required= {ins_req_other}')
                print(f'Total insert cost= Rs.{totIns_cost_other}')
                sheet1.write(3*pieces, 20, ins_req_other)
                sheet1.write(3*pieces, 22, totIns_cost_other)

            elif auto_input.upper() == 'M' :
                ins_req_other = float(input('Enter no. of inserts required= '))
                ins_cost_other = float(input('Enter price of 1 insert= Rs.'))
                totIns_cost_other = float(ins_req_other * ins_cost_other)
                act_totIns_cost.append(totIns_cost_other)
                print(f'Number of inserts required= {ins_req_other}')
                print(f'Total insert cost= Rs.{totIns_cost_other}')
                sheet1.write(3*pieces, 20, ins_req_other)
                sheet1.write(3*pieces, 22, totIns_cost_other)

            else:
                print('Invalid Input')

            

            totMC_cost_other = float(kw_cost_other + manP_cost_other + totIns_cost_other)
            print(f'Total production expense/setting= Rs.{totMC_cost_other}')

            act_totMC_cost.append(totMC_cost_other)
            break

     








    
act_totMC_cost = float(sum(costs for costs in act_totMC_cost))
miscost = input("Do you want to include miscellaneous cost? (Y/N) ")
act_Tmch_time = float(sum(time for time in act_Tmch_time))

if miscost.upper() == 'Y':
    mcdep = float(input('Enter M/C deppresiation rate per hour= Rs.'))
    mcdep = float(mcdep * act_Tmch_time)
    coolcost = float(input('Enter coolant cost= Rs.'))
    lubcost = float(input('Enter lubricant cost= Rs.'))
    anyothercost = float(input('Enter any other cost= Rs.'))
    
    miscost = float(coolcost + lubcost + anyothercost)

elif miscost.upper() == 'N':
    miscost = 0

else:
    print('Invalid option!')

sheet1.write(3*pieces, 23, miscost)

act_totMC_cost = float(act_totMC_cost + miscost)
print(f'Total production expense= Rs.{act_totMC_cost}')


exp_job = float(act_totMC_cost/des_qty)
print(f'Expense per job= Rs.{exp_job}')

sheet1.write(3*pieces, 24, act_totMC_cost)
sheet1.write(3*pieces, 25, exp_job)

offer_rate = float(input('Enter offer rate: Rs.'))
sp = float(offer_rate * des_qty)
print(f'Total S.P.= Rs.{sp}')

sheet1.write(3*pieces, 26, offer_rate)
sheet1.write(3*pieces, 27, sp)

chipcost_kg = float(input("Enter cost of chips/kg= Rs."))
chipcost = float(((mat_rmv*jobs*mat_den)/1000)*chipcost_kg)
print(f'Total S.P. of chips= Rs.{chipcost}')

profit = float(sp+chipcost-act_totMC_cost)
profitper = float(round((profit/act_totMC_cost)*100))

print(f'Total time required to finish machining the lot= {act_Tmch_time} hrs.')

hrs8_shift = float(act_Tmch_time / 8)
hrs10_shift = float(act_Tmch_time / 10)
hrs12_shift = float(act_Tmch_time / 12)

print(f'''
print(f'Total profit= Rs.{profit}')
print(f'Profit percentage= {profitper}%')

Number of 8 hours shifts required = {hrs8_shift}
Number of 10 hours shifts required= {hrs10_shift}
Number of 12 hours shifts required= {hrs12_shift}
''')

sheet1.write(3*pieces+2, 16, '8 hours shifts req.')
sheet1.write(3*pieces+2, 17, hrs8_shift)
sheet1.write(3*pieces+3, 16, '10 hours shifts req.')
sheet1.write(3*pieces+3, 17, hrs10_shift)
sheet1.write(3*pieces+4, 16, '12 hours shifts req.')
sheet1.write(3*pieces+4, 17, hrs12_shift)

sheet1.write(3*pieces, 28, profit)
sheet1.write(3*pieces, 29, profitper)
wb.save('CNC Job Quotation.xls')






plt.figure(1)
plt.xlabel('Feed/rev.(mm)')
plt.ylabel('Total Machining hrs.')
plt.plot(arbit_feed_rev, arbit_Tmch_time1)
plt.title('Feed vs. M/C Time')
plt.figure(2)
plt.xlabel('Feed/rev.(mm)')
plt.ylabel('Total Machining cost')
plt.plot(arbit_feed_rev, arbit_totMC_cost1)
plt.title('Feed vs. M/C Cost')
plt.figure(3)
plt.xlabel('Cutting speed(m/min)')
plt.ylabel('Total Machining hrs.')
plt.plot(arbit_vc, arbit_Tmch_time2)
plt.title('Cutting speed vs. M/C Time')
plt.figure(4)
plt.xlabel('Cutting speed(m/min)')
plt.ylabel('Total Machining cost')
plt.plot(arbit_vc, arbit_totMC_cost2)
plt.title('Cutting speed vs. M/C Cost')
plt.figure(5)
plt.xlabel('Cutting speed(m/min)')
plt.ylabel('Life of 1 insert(min)')
plt.plot(arbit_vc, arbit_ins_life2)
plt.title('Cutting speed vs. Insert Life')
plt.figure(6)
plt.xlabel('No. of jobs')
plt.ylabel('Total Machining hrs.')
plt.plot(arbit_jobs, arbit_Tmch_time3, label='Total Qty.')
plt.plot(arbit_rw_qty, arbit_Tmch_time4, label='Rework Qty.')
plt.title('No. of jobs vs. M/C Time')
plt.show()