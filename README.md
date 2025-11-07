# Coded UI Test External Plugin Source Project for WPF GridControl and WPF GridDataControl

This repository contains the coded ui test external plugin source project for [WPF GridControl](https://www.syncfusion.com/wpf-controls/excel-like-grid) and [WPF GridDataControl](https://help.syncfusion.com/wpf/classic/griddata/overview).

The Coded UI Test(CUIT) extension is a technology manager that tells Coded UI Test to use UI automation for Syncfusion WPF controls.

To test the Coded UI supported controls with CUITs, build the Extension project and place in the following mentioned location.

<table>
<tr>
<th>Controls</th>
<th>Compile assemblies</th>
<th>Adding Extension assembly</th>
<th>Extension Project</th>
</tr>
<tr>
<td>
GridControl
</td>
<td>
Grid.WPF
</td>
<td>
Syncfusion.VisualStudio.TestTools.UITest.GridExtension.dll
</td>
<td>
Get the Extension project from this repository.
</td>
</tr>
</table>

To run the CUITs, follow the steps:

1. Run the Extension project and build it.
2. You can get the following tabulated assembly from the bin folder.

The above assembly must be placed in the following directory based on your Visual Studio version.

<table>
<tr>
<th>Visual Studio Version</th>
<th>Relative Path</th>
</tr>
<tr>
<td>
2010
</td>
<td>
C:\Program Files (x86)\Common Files\Microsoft Shared\VSTT\10.0\UITestExtensionPackages
</td>
</tr>
<tr>
<td>
2015
</td>
<td>
C:\Program Files (x86)\Common Files\Microsoft Shared\VSTT\14.0\UITestExtensionPackages
</td>
</tr>
<tr>
<td>
2017
</td>
<td>
C:\Program Files (x86)\Common Files\Microsoft Shared\VSTT\15.0\UITestExtensionPackages
</td>
</tr>
</table>

The Extension package should be installed(For example, Syncfusion.VisualStudio.TestTools.UITest.SfGridExtension.dll) in GAC location. Refer to the MSDN link for [GAC](https://learn.microsoft.com/en-us/previous-versions/dotnet/netframework-2.0/ex0ss12c(v=vs.80)) installation.
