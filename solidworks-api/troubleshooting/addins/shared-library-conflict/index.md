---
layout: sw-addin-fix
title: How to fix the error of SOLIDWORKS add-ins sharing common libraries
caption: Add-ins which are using shared libraries cannot work together
description: Fixing the issue of using different versions of shared library by enabling binding redirect
labels: [add-in, troubleshooting, shared library]
---
## Symptoms

There are several SOLIDWORKS add-ins (usually from the same supplier) which cannot work together. SOLIDWORKS may crash or misbehave. Add-ins are working correctly if loaded independently.

## Cause

When same library (even of different versions) are used by different projects within the same application domain (e.g. add-in in SOLIDWORKS) .NET framework will use the cached library. The cached library will be the one which is accessed first. For example the library can be accessed when add-in button is clicked.

This results in the issues when library is not backward and forward compatible (i.e. version is supported by both newer and older applications). This is usually not the case for the libraries as behaviour may be changed, bugs fixed or regression issues introduced in the newer versions of library.

This introduces the possible conflicts when resolving the assembly references.

## Resolution

Sign conflicting assembly with a [strong name](https://docs.microsoft.com/en-us/dotnet/framework/app-domains/how-to-sign-an-assembly-with-a-strong-name). In this cases version specific assemblies will be used which will resolve conflict.

Hoverer, it might be the case where main project A refers the shared dll B with version 1 and also refers dll C which refers dll B with version 2, which means that it is required to have version 1 and 2 of B loaded at the same time. As dlls are usually compiled in the same directory it is either required to add them to different folders or use [Binding Redirect](https://docs.microsoft.com/en-us/dotnet/framework/configure-apps/file-schema/runtime/bindingredirect-element) element to redirect different versions of the shared library:

Add the following snippet to **app.config** file:

~~~ xml
<?xml version="1.0" encoding="utf-8" ?>
<configuration>
	<runtime>
		<assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
			<dependentAssembly>
				<assemblyIdentity name="[Assembly Name]" publicKeyToken="[Public Key Token]" culture="neutral" />
				<bindingRedirect oldVersion="0.0.0.0-9999.9999.9999.9999" newVersion="[Current Version]" />
			</dependentAssembly>
		</assemblyBinding>
	</runtime>
</configuration>
~~~

You can use the following snippet to find the required identity information (i.e. assembly name, version, public key token and culture) from the shared library.

~~~ cs
System.Diagnostics.Debug.Print(typeof([Any type from the shared assembly]).Assembly.FullName);
~~~

This will be printed as 

~~~
[Assembly Name], Version=[Version], Culture=[Culture], PublicKeyToken=[Public Key Token]
~~~

Video Demonstration: 

{% youtube { id: ZeWDoJ5TC7o } %}

Be aware of backward compatibility when using binding redirect, i.e. redirecting from version 1 to 2 requires backward compatibility, otherwise this solution will not work.

If shared assembly is not signed with a strong name it is possible to resolve the conflict at runtime by capturing the [AppDomain::AssemblyResolve](https://docs.microsoft.com/en-us/dotnet/api/system.appdomain.assemblyresolve?view=netframework-4.8) event and returning the resolved assembly from the method handler.
