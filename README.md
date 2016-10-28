# AnalysisTools Display
=========
### A Rhino 5 C++ plugin to display the results of analysis meshes in Rhino.

## Overview
The AnalysisTools plug-in is designed to import your analysis data into Rhino 5 for Windows for visual inspection using false color analysis. The plug-in can be used to read in result files from various analysis software to aid in further refinement of the model. The resulting analysis mesh in Rhino can be used to see 3D false color plots of velocity, pressure, solar radiance, and more.

The plugin currently reads the [.TP format](https://people.sc.fsu.edu/~jburkardt/data/tec/tec.html). See the [Samples directory](https://github.com/dalefugier/AnalysisTools/tree/master/Samples) for these files.  It is easy enough to modify this code to read other formats.

## More Information
* See the [AnalysisTools](http://www.food4rhino.com/app/analysistools) plug-in page on [Food4Rhino](http://www.food4rhino.com/) for detailed information about this plugin.
* Join the [Developer forums](http://discourse.mcneel.com/c/rhino-developer) to ask any questions about these tools.
* Email [devsupport@mcneel.com](mailto:devsupport@mcneel.com) for any input or detailed questiosn you have.


## Build

To build the sample, you are going to need:

* [Rhinoceros 5 for Windows](http://www.rhino3d.com)
* [Rhinoceros 5 C++ SDK](http://developer.rhino3d.com/guides/cpp/installing_tools_windows/)
* Microsoft Visual C++ 2010 (for 64-bit builds)
* Microsoft Visual C++ 2005 (optional, for 32-bit builds)

For detailed information on setting up and compiling a Rhino plugin see the [Creating your first C/C++ Plugin for Rhino Tutorial](http://developer.rhino3d.com/guides/cpp/your_first_plugin_windows/)
