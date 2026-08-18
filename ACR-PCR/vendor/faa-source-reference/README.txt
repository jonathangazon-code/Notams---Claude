FAA VB source, kept for reference only - NOT part of the build.

These two projects (ACClassLib, ICAOModels) are the FAA's own aircraft-library
code. They cannot be built in SharpDevelop: they are ToolsVersion 15.0 (Visual
Studio 2017) and use VB14 syntax (36 uses of NameOf) that SharpDevelop's VB
compiler does not accept.

They are not needed. AcrTool/AircraftLibrary.cs reads aircraft.xml directly, and
the five values it takes are read straight out of the same XML by clsAC.InitACLib
with no transformation (clsAC.vb lines 222-274):

    libCP        = Cp/us                 tyre pressure, psi
    libGL        = _GrossWeight/us       MTOW, lb
    libMGpcntPCN = MgPercentPCN          weight share on one truck pair
    libNWheels   = WheelCoordinates count
    libTX/libTY  = WheelCoordinates X/Y us, written 1-based

Keep them here to check that mapping, or if the geometry logic in modAC.vb is
ever needed for an aircraft type beyond the four this tool covers.
