from win32com import client
import os
from scipy.signal import find_peaks
import pandas as pd
from theoritical_plot import solve_equation
import math
import csv

class HFSS:
    def __init__(self, file_name="new"):
        self.oAnsoftApp = client.Dispatch('AnsoftHfss.HfssScriptInterface')
        self.oDesktop = self.oAnsoftApp.GetAppDesktop()
        if file_name == "new":
            self.oProject = self.oDesktop.NewProject()
        else:
            _base_path = os.getcwd()
            _path = os.path.join(_base_path, 'Prj{}.aedt'.format(file_name))
            self.oProject = self.oDesktop.OpenProject(_path)
        self.oProject.InsertDesign('HFSS', 'HFSSDesign1', 'DrivenModal1', '')
        self.oDesign = self.oProject.SetActiveDesign("HFSSDesign1")
        self.oEditor = self.oDesign.SetActiveEditor("3D Modeler")
        self.oModule = self.oDesign.GetModule('BoundarySetup')
        self.transparency = 0.5

    def set_variable(self, _var_name, _var_value):
        _NAME = 'NAME:' + _var_name
        _VALUE = str(_var_value) + 'mm'
        self.oDesign.ChangeProperty(["NAME:AllTabs",
                                     ["NAME:LocalVariableTab",
                                      ["NAME:PropServers", "LocalVariables"],
                                      ["NAME:NewProps",
                                       [_NAME, "PropType:=", "VariableProp", "UserDef:=", True, "Value:=", _VALUE]]]])

    def reset_var(self, _var_name, _var_value):
        _NAME = 'NAME:' + _var_name
        _VALUE = str(_var_value) + 'mm'
        self.oDesign.ChangeProperty(
            ["NAME:AllTabs",
             ["NAME:LocalVariableTab",
              ["NAME:PropServers","LocalVariables"],
              ["NAME:ChangedProps",
               [_NAME, "Value:="		, _VALUE]]]])

    def drawrectangle(self, x_p, y_p,z_p,width,height, name, _dir='Z',color="(255,128,0)"):
        self.oEditor.CreateRectangle(
            [
                "NAME:RectangleParameters",
                "IsCovered:="	, True,
                "XStart:="		, x_p,
                "YStart:="		, y_p,
                "ZStart:="		, z_p,
                "Width:="		, width,
                "Height:="		, height,
                "WhichAxis:="		, _dir
            ],
            [
                "NAME:Attributes",
                "Name:="		, name,
                "Flags:="		, "",
                "Color:="		, color,
                "Transparency:="	, 0,
                "PartCoordinateSystem:=", "Global",
                "UDMId:="		, "",
                "MaterialValue:="	, "\"vacuum\"",
                "SurfaceMaterialValue:=", "\"\"",
                "SolveInside:="		, True,
                "IsMaterialEditable:="	, True,
                "UseMaterialAppearance:=", False
            ])

    def drawbox(self,x_p,y_p,z_p,x_size,y_size,z_size,name,color="(143 175 143)"):
        self.oEditor.CreateBox(
            ["NAME:BoxParameters", "XPosition:=", x_p,"YPosition:=", y_p,"ZPosition:=", z_p,"XSize:=", x_size,
             "YSize:=", y_size,"ZSize:=", z_size],
            ["NAME:Attributes","Name:=", name,
             "Flags:=","",
             "Color:="		, color,
             "Transparency:="	, 0.3,
             "PartCoordinateSystem:=", "Global",
             "UDMId:="		, "",
             "MaterialValue:="	, "\"vacuum\"",
             "SurfaceMaterialValue:=", "\"\"",
             "SolveInside:="		, True,
             "IsMaterialEditable:="	, True,
             "UseMaterialAppearance:=", False
             ])

    def define_material(self, name, Er):
        _name = "NAME:" + name
        oDefinitionManager = self.oProject.GetDefinitionManager()
        oDefinitionManager.AddMaterial(
            [_name,
             "CoordinateSystemType:=", "Cartesian",
             "BulkOrSurfaceType:="	, 1,
             ["NAME:PhysicsTypes",
              "set:="			, ["Electromagnetic"]],
             "permittivity:="	, Er])

    def set_material(self, _obj, _mat='pec'):
        self.oEditor.AssignMaterial(
            [
                "NAME:Selections",
                "AllowRegionDependentPartSelectionForPMLCreation:=", True,
                "AllowRegionSelectionForPMLCreation:=", True,
                "Selections:="	, _obj
            ],
            [
                "NAME:Attributes",
                "MaterialValue:="	, "\"" + _mat + "\"",
                "SolveInside:="		, True,
                "IsMaterialEditable:="	, True,
                "UseMaterialAppearance:=", False
        ])

    def openRegion(self, frequency="6.5GHz"):
        oModule = self.oDesign.GetModule("ModelSetup")
        oModule.CreateOpenRegion(
            ["NAME:Settings",
             "OpFreq:="		, frequency,
             "Boundary:="		, "Radiation",
             "ApplyInfiniteGP:="	, False
             ])

    def plane_wave(self):
        oModule = self.oDesign.GetModule("BoundarySetup")
        oModule.AssignPlaneWave(
            ["NAME:IncPWave1",
             "IsCartesian:="		, True,
             "EoX:="			, "0",
             "EoY:="			, "0",
             "EoZ:="			, "1",
             "kX:="			, "-1",
             "kY:="			, "0",
             "kZ:="			, "0",
             "OriginX:="		, "0mm",
             "OriginY:="		, "0mm",
             "OriginZ:="		, "0mm",
             "IsPropagating:="	, True,
             "IsEvanescent:="	, False,
             "IsEllipticallyPolarized:=", False])

    def editwave(self,eox,eoy,eoz,kx,ky,kz):
        oModule = self.oDesign.GetModule("BoundarySetup")
        oModule.EditIncidentWave("IncPWave1",
                                 ["NAME:IncPWave1",
                                  "IsCartesian:="		, True,
                                  "EoX:="			, str(eox),
                                  "EoY:="			, str(eoy),
                                  "EoZ:="			, str(eoz),
                                  "kX:="			, str(kx),
                                  "kY:="			, str(ky),
                                  "kZ:="			,str(kz),
                                  "OriginX:="		, "0mm",
                                  "OriginY:="		, "0mm",
                                  "OriginZ:="		, "0mm",
                                  "IsPropagating:="	, True,
                                  "IsEvanescent:="	, False,
                                  "IsEllipticallyPolarized:=", False])

    def set_up(self, frequency, name ="Setup1"):
        _Name = "NAME:"+name
        oModule = self.oDesign.GetModule("AnalysisSetup")
        oModule.InsertSetup("HfssDriven",
                            ["NAME:Setup1",
                             "AdaptMultipleFreqs:="	, False,
                             "Frequency:="		, frequency,
                             "MaxDeltaE:="		, 0.02,
                             "MaximumPasses:="	, 20,
                             "MinimumPasses:="	, 1,
                             "MinimumConvergedPasses:=", 1,
                             "PercentRefinement:="	, 30,
                             "IsEnabled:="		, True,
                             "BasisOrder:="		, 1,
                             "DoLambdaRefine:="	, True,
                             "DoMaterialLambda:="	, True,
                             "SetLambdaTarget:="	, False,
                             "Target:="		, 0.3333,
                             "UseMaxTetIncrease:="	, False,
                             "UseDomains:="		, False,
                             "UseIterativeSolver:="	, False,
                             "SaveRadFieldsOnly:="	, False,
                             "SaveAnyFields:="	, True,
                             "IESolverType:="	, "Auto",
                             "LambdaTargetForIESolver:=", 0.15,
                             "UseDefaultLambdaTgtForIESolver:=", True])

    def sweep(self,start,end):
        oModule = self.oDesign.GetModule("AnalysisSetup")
        oModule.InsertFrequencySweep("Setup1",
                                     ["NAME:Sweep",
                                      "IsEnabled:="		, True,
                                      "RangeType:="		, "LinearStep",
                                      "RangeStart:="		, start,
                                      "RangeEnd:="		, end,
                                      "RangeStep:="		, "0.01GHz",
                                      "Type:="		, "Discrete",
                                      "SaveFields:="		, True,
                                      "SaveRadFields:="	, False,
                                      "ExtrapToDC:="		, False])

    def infiniteSph(self):
        oModule = self.oDesign.GetModule("RadField")
        oModule.InsertFarFieldSphereSetup(
            ["NAME:Infinite Sphere1",
             "UseCustomRadiationSurface:=", False,
             "ThetaStart:="		, "-180deg",
             "ThetaStop:="		, "180deg",
             "ThetaStep:="		, "45deg",
             "PhiStart:="		, "0deg",
             "PhiStop:="		, "360deg",
             "PhiStep:="		, "15deg",
             "UseLocalCS:="		, False])

    def creat_result(self, eox,eoy,eoz,kx,ky,kz,phi, theta,path):
        oModule = self.oDesign.GetModule("ReportSetup")
        _phi = str(phi)+"deg"
        _theta = str(theta)+"deg"
        oModule.CreateReport("Normalized Bistatic RCS Plot 1", "Far Fields", "Rectangular Plot", "Setup1 : Sweep",
                             ["Context:="		, "Infinite Sphere1"],
                             ["Freq:="		, ["All"],
                              "Phi:="			, [_phi],
                              "Theta:="		, [_theta],
                              "dr_x:="		, ["Nominal"],
                              "dr_y:="		, ["Nominal"],
                              "dr_z:="		, ["Nominal"]],
                             ["X Component:="		, "Freq",
                              "Y Component:="		, ["dB(NormRCSTotal)"]], [])

        '''store csv file in certain location'''
        file_name = "Normalized_Bistatic_RCS_e_"+str(eox)+str(eoy)+str(eoz)+"_phi_"+str(phi)+"_theta_"+str(theta)+".csv"
        _base_path = os.getcwd()
        #_path = os.path.join(_base_path,"rcs_plot_diff_direction")
        #_path = os.path.join(_base_path,"whole_data_rcs_plot")
        _path = os.path.join(path,file_name)
        oModule.ExportToFile("Normalized Bistatic RCS Plot 1", _path)

        dataset = pd.read_csv(_path).to_numpy()
        peaks, _ = find_peaks(dataset[:,1], height=-30, threshold=0.001, prominence=2.5)

        frequency = dataset[peaks,0]
        return frequency

    def integrate(self,frequency,a,b,d,phase,path):
        _frequency = str(frequency)+"GHz"
        oModule = self.oDesign.GetModule("FieldsReporter")
        oModule.CalcStack("clear")
        oModule.CopyNamedExprToStack("Mag_E")
        oModule.EnterVol("dr")
        oModule.CalcOp("Integrate")
        oModule.ClcEval("Setup1 : Sweep",
                        ["Freq:="		, _frequency,
                         "Phase:="		, str(phase)+"deg",
                         "dr_x:="		, str(a)+"mm",
                         "dr_y:="		, str(b)+"mm",
                         "dr_z:="		, str(d)+"mm"])
        oModule.CalculatorWrite(path,
                                ["Solution:="		, "Setup1 : Sweep"],
                                ["Freq:="		, _frequency,
                                 "Phase:="		, str(phase)+"deg",
                                 "dr_x:="		, str(a)+"mm",
                                 "dr_y:="		, str(b)+"mm",
                                 "dr_z:="		, str(d)+"mm"])
        with open(path) as f:
            f_csv = csv.reader(f)
            headers_1 = next(f_csv)
            for row in f_csv:
                value = float(row[0])
        return value

    def export(self,frequency,dr_x,dr_y,dr_z,base_path,phase,type):
        _Frequency = str(frequency)+"GHz"
        x = str(dr_x)+"mm"
        y = str(dr_y)+"mm"
        z = str(dr_z)+"mm"

        end_x = str(dr_x/2)+"mm"
        start_x = "-"+end_x
        end_y = str(dr_y/2)+"mm"
        start_y = "-"+end_y
        end_z = str(dr_z/2)+"mm"
        start_z = "-"+end_z
        step_x = str(dr_x/20)+"mm"
        step_y = str(dr_y/20)+"mm"
        step_z = str(dr_z/20)+"mm"
        oModule = self.oDesign.GetModule("FieldsReporter")
        oModule.CalcStack("clear")
        ###############################################################
        oModule.CopyNamedExprToStack(type)
        #oModule.CopyNamedExprToStack("Vector_E")
        #oModule.CopyNamedExprToStack("ComplexMag_E")

        _path = os.path.join(base_path,'datatype_{}.fld'.format(type))
        oModule.ExportOnGrid(_path, [start_x, start_y, start_z], [end_x, end_y, end_z], [step_x, step_y, step_z], "Setup1 : Sweep",
                             ["Freq:="		, _Frequency,
                              "Phase:="		, str(phase)+"deg",
                              "dr_x:="		, x,
                              "dr_y:="		, y,
                              "dr_z:="		, z
                              ], True, "Cartesian", ["0mm", "0mm", "0mm"], False)


    def save_prj(self):
        _base_path = os.getcwd()
        _prj_num = 1
        '''
        while True:
            _path = os.path.join(_base_path, 'Prj{}.aedt'.format(_prj_num))
            if os.path.exists(_path):
                _prj_num += 1
            else:
                break
        '''
        _path = os.path.join(_base_path, 'Prj{}.aedt'.format(_prj_num))
        self.oProject.SaveAs(_path, True)

    def run(self):
        self.oDesktop.RestoreWindow()
        self.oDesign.Analyze('Setup1')

    def release(self):
        #self.save_prj()
        self.oDesktop.CloseAllWindows()
        self.oDesktop.QuitApplication()
        del self.oEditor
        del self.oDesign
        del self.oProject
        del self.oDesktop
        del self.oAnsoftApp


def eval_boundary_Point(a,b,d,er,m,n,p):
    #eval_k0 = 20*math.pi
    a = a/1000
    b = b/1000
    d = d/1000
    a_norm = a / m
    b_norm = b /n
    d_norm = d /p
    kx = math.pi / a_norm
    ky = math.pi / b_norm
    kz = math.pi / d_norm
    eval_k0 = math.sqrt((kx**2 + ky**2 + 0.75*kz**2)/er)
    kz_root,k0_root_1 = solve_equation(a,b,d,er,kz*0.8,eval_k0)
    eval_k0 = math.sqrt((kx**2 + kz**2 + 0.75*ky**2)/er)
    ky_root,k0_root_2 = solve_equation(a,d,b,er,ky*0.8,eval_k0)
    eval_k0 = math.sqrt((kz**2 + ky**2 + 0.75*kx**2)/er)
    kx_root,k0_root_3 = solve_equation(d,b,a,er,kx*0.8,eval_k0)
    print(k0_root_1,k0_root_2,k0_root_3)
    return k0_root_1,k0_root_2,k0_root_3
    #start_point = min(k0_root_1,k0_root_2,k0_root_3)
    #start_point = 3*start_point/20/math.pi
    #eval_k0 = 60*math.pi


def creat_Model():
    #create folders
    _base_path = os.getcwd()
    _e_path = os.path.join(_base_path, 'TM_test')
    if not os.path.exists(_e_path):
        os.mkdir(_e_path)
    _rcs_path = os.path.join(_base_path,"TM_rcs")
    _phase_path = os.path.join(_base_path,'phase.fld')
    if not os.path.exists(_rcs_path):
        os.mkdir(_rcs_path)
    #x = 9.5
    #y = 13.5
    #z = 5.5
    x = 13.5
    y = 14.175
    z = 5.5
    er = 45

    for w in range(1):
        for g in range(1):

            #width = round(y*(0.9+0.05*w),3)
            #height = round(z*(0.9+0.05*g),3)
            #dimension_folder_name = "width_"+str(width)+"_height_"+str(height)
            '''
            if w < 2:
                length = round(x*(0.9+0.05*w),3)
            else:
                length = round(x*(0.95+0.05*w),3)
            '''
            length = x
            #width = round(y*(0.9+0.05*w),3)
            width = y
            height = z

            dimension_folder_name = "width_"+str(width)+"_fixed_height_"+str(height)
            #dimension_folder_name = "length_"+str(length)

            if not os.path.exists(os.path.join(_e_path,dimension_folder_name)):
                os.mkdir(os.path.join(_e_path,dimension_folder_name))
            if not os.path.exists(os.path.join(_rcs_path,dimension_folder_name)):
                os.mkdir(os.path.join(_rcs_path,dimension_folder_name))

            s_1,s_2,s_3 = eval_boundary_Point(length,width,height,er,1,1,1)
            start_Point = min(s_1,s_2,s_3)
            start_Point = 3*start_Point/20/math.pi

            e_1,e_2,e_3 = eval_boundary_Point(length,width,height,er,2,1,1)
            end_Point = max(e_1,e_2,e_3)
            e_1,e_2,e_3 = eval_boundary_Point(length,width,height,er,1,2,1)
            end_Point = max(e_1,e_2,e_3,end_Point)
            e_1,e_2,e_3 = eval_boundary_Point(length,width,height,er,1,1,2)
            end_Point = max(e_1,e_2,e_3,end_Point)
            end_Point = 3*end_Point/20/math.pi

            start_Point = start_Point - 0.5
            start_Point = max(start_Point,0.1)
            start_Point = round(start_Point,3)
            end_Point = end_Point + 0.5
            end_Point = round(end_Point,3)
            fr = (end_Point + start_Point)/2
            fr = str(fr)+"GHz"
            start_Point = str(start_Point)+"GHz"
            end_Point = str(end_Point)+"GHz"
            print(start_Point)
            print(end_Point)

            try:
                h
            except NameError:
                print("No h object, Creating...")
                h = HFSS()
            else:
                print("Delete h, recreating...")
                h.release()
                del h
                h = HFSS()

            h.set_variable("dr_x",length)
            h.set_variable("dr_y",width)
            h.set_variable("dr_z",height)
            h.drawbox("-"+"dr_x"+"/2","-"+"dr_y"+"/2","-"+"dr_z"+"/2","dr_x","dr_y","dr_z","dr")
            h.openRegion(start_Point)
            h.define_material("DRA_45","45")
            h.set_material("dr","DRA_45")
            h.plane_wave()
            h.set_up(fr)
            h.infiniteSph()
            h.sweep(start_Point,end_Point)
            print("dimension:",length,width,height,start_Point,end_Point)

            ##########change excitation direction
            '''
            for index in range(1):
                kx,ky,kz,eox,eoy,eoz,phi,theta = 0,1,0,0,0,1,90,90


                print(kx,ky,kz,eox,eoy,eoz,phi, theta)
                h.editwave(eox,eoy,eoz,kx,ky,kz)
                h.run()
                
                frequency = h.creat_result(eox,eoy,eoz,kx,ky,kz,phi,theta,os.path.join(_rcs_path,dimension_folder_name))
                
                _field_path = os.path.join(_e_path,dimension_folder_name,"e_"+str(eox)+str(eoy)+str(eoz)+"_phi_"+str(phi)+"_theta_"+str(theta))
                os.mkdir(_field_path)
                try:
                    for f in range(len(frequency)):
                        original_phase = 0
                        original_mag = 0
                        for p in range(36):
                            phase = p*10
                            mag = h.integrate(frequency[f],x,width,height,phase,_phase_path)
                            if mag > original_mag:
                                original_mag = mag
                                original_phase = phase
                            #read mag data & compare & export in certain phase
                        norm_frequency = frequency[f]
                        norm_frequency = round(norm_frequency,3)
                        whole_data_path = os.path.join(_field_path,"phase{}_peak{}_frequency{}".format(original_phase,f,norm_frequency))
                        os.mkdir(whole_data_path)
                        h.export(frequency[f],x,width,height,whole_data_path,original_phase,"Mag_E")
                        #h.export(frequency[f],x,y,z,whole_data_path,original_phase,"ComplexMag_E")
                        h.export(frequency[f],x,width,height,whole_data_path,original_phase,"Vector_E")
                    #print("e_"+str(eox)+str(eoy)+str(eoz)+"_phi_"+str(phi)+"_theta"+str(theta))
                except IOError:
                    print("Cannot export")
            '''


            for index in range(6):
                if index == 3:
                    h.release()
                    del h
                    h = HFSS()
                    h.set_variable("dr_x",length)
                    h.set_variable("dr_y",width)
                    h.set_variable("dr_z",height)
                    h.drawbox("-"+"dr_x"+"/2","-"+"dr_y"+"/2","-"+"dr_z"+"/2","dr_x","dr_y","dr_z","dr")
                    h.openRegion(start_Point)
                    h.define_material("DRA_45","45")
                    h.set_material("dr","DRA_45")
                    h.plane_wave()
                    h.set_up(fr)
                    h.infiniteSph()
                    h.sweep(start_Point,end_Point)
                kx,ky,kz,eox,eoy,eoz= 0,0,0,0,0,0
                theta = 90
                phi = 0
                if index / 2 < 1:
                    kx = 1
                    if index % 2 < 1:
                        eoy = 1
                    else:
                        eoz = 1
                else:
                    if index /2 <2:
                        ky = 1
                        phi = phi + 90
                        if index % 2 < 1:
                            eox = 1
                        else:
                            eoz = 1
                    else:
                        kz = 1
                        phi = 0
                        theta = 0
                        if index % 2 < 1:
                            eox = 1
                        else:
                            eoy = 1
                print(kx,ky,kz,eox,eoy,eoz,phi, theta)

                h.editwave(eox,eoy,eoz,kx,ky,kz)
                h.run()
                frequency = h.creat_result(eox,eoy,eoz,kx,ky,kz,phi,theta,os.path.join(_rcs_path,dimension_folder_name))
                _field_path = os.path.join(_e_path,dimension_folder_name,"e_"+str(eox)+str(eoy)+str(eoz)+"_phi_"+str(phi)+"_theta_"+str(theta))
                os.mkdir(_field_path)
                try:
                    for f in range(len(frequency)):
                        original_phase = 0
                        original_mag = 0
                        for p in range(36):
                            phase = p*10
                            mag = h.integrate(frequency[f],length,width,height,phase,_phase_path)
                            if mag > original_mag:
                                original_mag = mag
                                original_phase = phase
                            #read mag data & compare & export in certain phase
                        norm_frequency = frequency[f]
                        norm_frequency = round(norm_frequency,3)
                        whole_data_path = os.path.join(_field_path,"phase{}_peak{}_frequency{}".format(original_phase,f,norm_frequency))
                        os.mkdir(whole_data_path)
                        h.export(frequency[f],length,width,height,whole_data_path,original_phase,"Mag_E")
                        #h.export(frequency[f],x,y,z,whole_data_path,original_phase,"ComplexMag_E")
                        h.export(frequency[f],length,width,height,whole_data_path,original_phase,"Vector_E")
                        h.export(frequency[f],length,width,height,whole_data_path,original_phase,"Mag_H")
                        h.export(frequency[f],length,width,height,whole_data_path,original_phase,"Vector_H")
                    #print("e_"+str(eox)+str(eoy)+str(eoz)+"_phi_"+str(phi)+"_theta"+str(theta))
                except IOError:
                    print("Cannot export")



if __name__ == '__main__':
    creat_Model()
