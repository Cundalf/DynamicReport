# Dynamic Report (CADR)
CADR (Cundari Agustin Dynamic Report) es un proyecto desarrollado en Visual Basic 6 e Intersystem Cache. Lo desarrolle con el fin de evitar generar compilaciones de la aplicación y generar reportes desde la base de datos. Vale aclarar que tenia una alta demanda de reportes en ese momento. Espero que el que use este código, le sea tan útil como a mi  cuando lo desarrolle.

*English:*

CADR (Cundari Agustin Dynamic Report) is a project developed in Visual Basic 6 and Intersystem Cache. I developed it in order to avoid generating compilations of the application and generating reports from the database. He was in high demand for reports at the time. I hope this code is as useful to those who need it as it was to me when I developed it.

# Características / Features
 - Generar Reportes dinamicamente en Visual Basic 6 (Menu Item y Form).
 - Exportar los datos  de un ListView en Microsoft Excel, Libreoffice Calc y CSV.

*English:*
- Generate reports dynamically in Visual Basic 6 (Menu Item and Form).
- Export the data of a ListView in Microsoft Excel, Libreoffice Calc and CSV.

# Dependencias / Dependencies
Dentro del proyecto encontraran todo lo necesario para poder ejecutar y probar la aplicación. Deben tener en cuenta que este proyecto simplemente es una demostración de reportes dinámicos que pueden adaptar a sus proyectos.

Pueden registrar los archivos OCX y DLL con el script **RegOcxDll.bat** ubicado en /Dependencies/RegOcxDll.bat

*English:*

Within the project you will find everything you need to be able to run and test the application. Please note that this project is simply a demonstration of dynamic reporting that you can adapt to your projects.

You can register the OCX and DLL files with the script **RegOcxDll.bat** located at /Dependencies/RegOcxDll.bat

## VB6 Referencias / References
- [Default] Visual Basic For Applications (msvbvm60.dll)
- [Default] Visual Basic runtime objects and procedures (msvbvm60.dll)
- [Default] Visual Basic objects and procedures (VB6.OLB)
- [Default] OLE Automation (stdole2.tlb)
- Microsoft Scripting Runtime (scrrun.dll)
 
## VB6 Componentes / Components
- Microsoft Windows Common Controls 6.0 (SP6) (MSCOMCTL.OCX)
- Microsoft Windows Common Controls-2 6.0 (SP6) (MSCOMCT2.OCX)
- VisM 7.1 ActiveX Control module (VISM.OCX)

## Intersystem Cache
Dentro de la carpeta **Database** encontrara el archivo CADR.MAC que contiene las funciones de Cache necesarias para que el proyecto funcione. Tambien, encontrara el archivo INSTALLCADR.MAC el cual contiene un ejemplo de los globales necesarios para que el proyecto pueda generar algunos menus.

> Recuerde: Una vez incorporado ambos archivos debe ejecutar la rutina.
>
> D ^INSTALLCADR

*English:*

Inside the **Database** folder you will find the CADR.MAC file that contains the Cache functions necessary for the project to work. Also, you will find the INSTALLCADR.MAC file which contains an example of the global ones necessary for the project to generate some menus.

> Remember: Once both files are incorporated, you must execute the routine.
>
> D ^INSTALLCADR
