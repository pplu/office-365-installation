# Installing old versions of Office 365

There are cases you may wish to install an old version of your Office 365 applications (Excel, Word, Powerpoint). 

You may find yourself wanting to do this because some automation breaks on you because of an Office update, for example. 
The last time I found this was that in order to migrate data off of a Windows 2008 Server, Excel needed to be working 
correctly. The bad news was that downloading the Office 365 applications installer from my Office 365 subscription would 
not work, since Windows Server 2008 had fallen out of support some time ago.

Luckily, Microsoft can help you :)

# Office Deployment Tool

Welcome [https://www.microsoft.com/en-us/download/details.aspx?id=49117](ODT) (Office Deployment Tool). This utility
helps us download and install any version of Office 365 that has been published. Although its built to help automate 
Office 365 installations, if correctly instructed, we can control the exact version of Office applications that it 
installs. See an overview of its functionality [https://docs.microsoft.com/en-us/deployoffice/overview-office-deployment-tool](here).

All we need to do is to [https://www.microsoft.com/en-us/download/details.aspx?id=49117](download ODT), extract it to 
a directory (`c:\ODT` in this article), and write a configuration file:

# Configuring ODT

ODT is controlled with an XML file. Create a file `office.xml` in `c:\ODT\`:

```
<Configuration>
  <Add Version="16.0.7571.2109" OfficeClientEdition="64">
    <Product ID="O365BusinessRetail">
      <Language ID="en-us" />
    </Product>
  </Add>  
</Configuration>
```

Set the `Version` attribute according to the version you want installed.
Office 365 applications seem to be version `16.0`, followed by a *build number* (which you can find on the 
[https://docs.microsoft.com/en-gb/officeupdates/update-history-microsoft365-apps-by-date](version history page).
Note also that this config defaults to the *Current* channel. Take a look at `Channel` attribute, which is documented in
the [https://docs.microsoft.com/en-us/deployoffice/office-deployment-tool-configuration-options](configuration options of ODT).

# Download and install

ODT has different execution modes for preparing or executing different phases of an Office 365 automatic installation. First 
step is to download the Office installer.

Open a `cmd.exe` and `cd c:\ODT`:

```
setup.exe /download office.xml
```

You will see that in `C:\ODT\` that version of Office has been downloaded. Now to install it:

```
setup.exe /configure office.xml
```

Once the installation finishes, the configured build of Office 365 will be available

