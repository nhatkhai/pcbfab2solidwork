This script intended to generate Solidwork 3D model of the PCB from PCB Fab
drawing files (Drill, Gerber, 3D-BOM...) or KiCad PCBNew file (.brd file
only).

# History & HowItWork
  * ``macros/kicad_to_solidwork`` \
  This script were first developed for generated 3D model from KiCad 4.0.0
  version where PCBNew were using .brd file. The script was reading
  following layers:
    * PCB Edge Layer: Line, Arc, Circle for constructing the base PCB
      Outline, and pickup drilling hold infomation.
    * Top and Bottom Silk Layer: Line/DS, Circle/DC, T0 (Reference), T1
      (Value), and Text objects for reconstruct silk screen layer
    * Top and Bottom Cooper Layer: Text objects for reconstruct onto PCB
    * Component location, angle, and it's 3D models. The VRML file will
      also coarsely converted into Solidwork native format.

  * ``macros/gerber_to_solidwork`` \
  This script were developed after realization that KiCad changed their
  file formats.There was also a need to created Solidwork model for PCB
  created from difference programs. So choosing a set of standard
  fabrication file that most PCB software generated to send to PCB
  manufacture. These files were seem to be Gerber, Drill, Component
  position, and BOM files. However, 3D-BOM file would be more useful for
  place the 3D model of component into Solidwork. The VRML 3D file will
  coarsely converted into Solidwork native model.

  * All to conversion start with sketch of a interesting layer or objects.
    Then extruded them into a 3D shape/feature. Also, because of this, it
    is hard to create a sketch from Copper layer that can be extruded
    without Solidwork complaints. 

# Intruction to import script into Solidwork
  * Solidwork -> Tools -> Macro -> New..
  * Chose macro name of your choice, and press on Save
  * MS Visual Basic -> File -> Import File.. (Ctrl+M shortcut) to import
    all files in macros/libs, and either macros/gerber_to_solidwork or
    macros/kicad_to_solidwork files.
  * Now you should able to run macro from Solidwork -> Tools -> Macro ->
    Run menu

  [![Board3_Demo](https://img.youtube.com/vi/Xe5iQdEkxaU/0.jpg)](https://youtu.be/Xe5iQdEkxaU)

# 3D BOM Format
  * It take a CSV file as a table with following header:
    ```
    References, any-,  Scale , Offset , Rotation, 3DModel File Path
              ,thing, X, Y, Z, X, Y, Z, X, Y, Z ,(STEP; VMRL; SLDPRT)
    ```
    Path to 3D Model file can be:
      * Relative path to where the CSV file located
      * file extension are optional, can be left out

# Know Issues
  * If Gerber for PCB Edge did not enclosed, or have repeated drawing
    overlap may make Solidwork not happy to extrue sktech into feature.
    When this happend, the script will save the sketch, and continue on.
    The saved sketch, can be manually correct/edited and extruded.

