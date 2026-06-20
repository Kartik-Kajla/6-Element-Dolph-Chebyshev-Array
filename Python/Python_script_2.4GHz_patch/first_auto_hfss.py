import HFSS_auto.Geometry as geo
import HFSS_auto.Boundries as bd
import HFSS_auto.Excitation as ex
import HFSS_auto.Boolean as bl
import HFSS_auto.Edit as ed
import HFSS_auto.Setup as st
import HFSS_auto.Modal_Solution_Report as report
import HFSS_auto.Export_file as exp

import win32com.client

oAnsoftApp = win32com.client.Dispatch("Ansoft.ElectronicsDesktop")
oDesktop = oAnsoftApp.GetAppDesktop()
oProject = oDesktop.SetActiveProject("Python_auto")
oDesign = oProject.SetActiveDesign("patch_antenna")
oModule = oDesign.GetModule("BoundarySetup")
oModel = oDesign.GetModule("ModelSetup")
oAnalysis = oDesign.GetModule("AnalysisSetup")
oReport = oDesign.GetModule("ReportSetup")

# Open existing project

oProject = oDesktop.GetActiveProject()
oDesign = oProject.GetActiveDesign()
oEditor = oDesign.SetActiveEditor("3D Modeler")

geo.create_box(oEditor,-40, -40, 0,80, 80, 1.6, name = "Substrate",
               material="Rogers RT/duroid 5880 (tm)")
geo.create_rectangle(oEditor, -40, -40, 0, 80, 80, name = "Ground_Plane",
                     color = "(255 128 64)"
                     )

                     
bd.assign_finite_conductivity(oModule, "Ground_Plane", material = "copper",
                              name = "Finite_conductivity")




geo.create_rectangle(oEditor, -2.5, 40, 0, 1.6, 5, name="Lumped_port",axis="y")
ex.assign_lumped_port(oModule, "Lumped_port", [0,40,0], [0,40,1.6])

st.create_open_region(oModel, "2.4GHz")


i=6
length = 41.30
original = 41.22

st.create_setup(oAnalysis,setup_name=f"setup{i}",start_freq="0GHz",end_freq="10GHz",points=501,
                    sweep_type="Interpolating",solver_setting="Higher Speed")
    
geo.create_rectangle(oEditor, -3.5, 6, 1.6, 7, 20.66-6,name="Rect",color="(255 128 64)")
geo.create_rectangle(oEditor, -24.19, -length/2, 1.6, 49.38, 40.5, name = "Patch_Antenna",
                         color = "(255 128 64)")
                         
bd.assign_finite_conductivity(oModule, "Patch_Antenna", material = "copper",
                                  name = "Finite_conductivity_6")
bl.subtract(oEditor, "Patch_Antenna", "Rect")
geo.create_rectangle(oEditor, -2.5, i, 1.6, 5, 20+(20-i), name="Feed",color="(255 128 64)")
    
bd.assign_finite_conductivity(oModule, "Feed", material = "copper",
                                  name = "Finite_conductivity_7")

bl.unite(oEditor, "Patch_Antenna,Feed")
oDesign.Analyze(f"setup{i}")

report.create_sparameter_plot(oReport, "S_param_6",setup_name="setup6")
report.create_sparameter_table(oReport, "S_param_table_6",setup_name="setup6")
exp.export_report_to_csv(oReport, "S_param_table_6",
                            "C:/Users/Kartik/Desktop/ECE_SOFTWARE_wORKS/HFSS_Works/Exported_files/S Parameter Table_6.csv")
    
    
    
  



