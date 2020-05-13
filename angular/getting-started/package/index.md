---
layout: article
title: Angular package.json file overview
caption: Package File Overview
description: Overview of a package.json file to manage Angular project
lang: en
image: /angular/getting-started/package/angular-package.png
labels: [angular,package,json,dependencies,devDependencies,scripts]
order: 2
---
{% include img.html src="angular-package.png" width=400 alt="Angular package file" align="center" %}

Once you create new Angular application, you will see package.json file among the newly created files and folders. package.json file locates in project root and contains information about your web application. The main purpose of the file comes from its name *package*, so it'll contain the information about npm packages installed for the project.<br>

Let's take a look at main sections in package.json file.

## Project Metadata

Metadata contains information about your application.

~~~
  "name": "my-first-angular-app",
  "version": "0.0.0",
  "private": true,
~~~

Field         | Value                          | Description
--------------|--------------------------------|-----------
"name" | "my-first-angular-app"     | Your project name
"version" | "0.0.1" | Your project version
"private" | true | Project is private and can't be published in npm

You can add the following fields and values:

Field         | Value                          | Description
--------------|--------------------------------|-----------
"description" | "Some project description"     |
"main"        | "src/main.ts"                  | Entry point in the app. "src/main.ts" is default value for Angular application
"author"      | "Name, name@email.com, https://name.com"<br>or<br>"author": { <br> "name":"Name",<br>"email":"name@email.com",<br>"url":"https://name.com"} | Set all in one plain string <br> or <br> using the structure
"contributors"| ["Contributor, contributor@email.com, https://contributor.com"]<br>or<br>"contributors": [{<br>"name": "Contributor",<br>"email": "contributor@email.com",<br>"url": "https://contributor.com"<br>}]<br>  | You can add contributors information as string array<br>or<br>using the structure array
"bugs"        | "https://github.com/UserName/my-first-angular-app/issues"  | Link to bug tracking system, if any
"homepage"    | "https://site-name.com"                               | Link to homepage
"license"     | "MIT"                                                 | License of the application
"repository"  | "{"type": "git",<br> "url": "https://github.com/UserName/my-first-angular-app.git"}"                   | Repository location

## Scripts

This section describes Node scripts you can run in your application. As the code sample uses Angular CLI, all scripts are calling it.

~~~
  "scripts": {
    "ng": "ng",
    "start": "ng serve",
    "build": "ng build",
    "test": "ng test",
    "lint": "ng lint",
    "e2e": "ng e2e"
  },
~~~

You can put any cmd command in the script and you will be able to run it with npm. To run the script, just run it in command line from project location:

~~~
> npm start
~~~

This line with run the **ng serve** for you, which means start the application. You can clean existing and add your own new scripts.

## Dependencies

The list of packages installed as dependencies for this project are required at runtime.

~~~
"dependencies": {
  "@angular/animations": "~8.0.1",
  "@angular/common": "~8.0.1",
  "@angular/compiler": "~8.0.1",
  "@angular/core": "~8.0.1",
  "@angular/forms": "~8.0.1",
  "@angular/platform-browser": "~8.0.1",
  "@angular/platform-browser-dynamic": "~8.0.1",
  "@angular/router": "~8.0.1",
  "rxjs": "~6.4.0",
  "tslib": "^1.9.0",
  "zone.js": "~0.9.1"
},
~~~

Where **@angular/animations** is package name and **~8.0.0** is package version. You may notice that package versions description vary. The symbol in front of version says to npm install which package version to use<br>
**1.0.0** means strictly 1.0.0 version of the package<br>
**~8.0.0** means, 8.0.0 version or it's later patch version (third number may vary): 8.0, 8.0.x<br>
**^1.9.0** means, 1.9.0 version or it's later minor version (second number may vary): 1, 1.x <br>

Other version control symbols:

Symbol         | Description      
---------------|------------------
*  | Latest version of the package may be installed, including major version update (first number may vary)
>  | Version higher then specified should be installed
>= | Same or higher package version should be installed
<  | Version lower then specified should be installed
<= | Same or lower package version should be installed 

To install a new package run:

~~~
> npm i new-package-name-to-install
~~~

Where **i** is short for install and you will need to replace **new-package-name-to-install** with actual package name and it's version. Check the [https://www.npmjs.com/](https://www.npmjs.com/) for available packages. The line with package name and it's version will be added to "dependencies" list automatically after installation.

## Development Dependencies

The list of packages that are required only for development. This packages are installed only on developer's machine and will not be run for production build.

~~~
"devDependencies": {
  "@angular-devkit/build-angular": "~0.800.0",
  "@angular/cli": "~8.0.3",
  "@angular/compiler-cli": "~8.0.1",
  "@angular/language-service": "~8.0.1",
  "@types/node": "~8.9.4",
  "@types/jasmine": "~3.3.8",
  "@types/jasminewd2": "~2.0.3",
  "codelyzer": "^5.0.0",
  "jasmine-core": "~3.4.0",
  "jasmine-spec-reporter": "~4.2.1",
  "karma": "~4.1.0",
  "karma-chrome-launcher": "~2.2.0",
  "karma-coverage-istanbul-reporter": "~2.0.1",
  "karma-jasmine": "~2.0.1",
  "karma-jasmine-html-reporter": "^1.4.0",
  "protractor": "~5.4.0",
  "ts-node": "~7.0.0",
  "tslint": "~5.15.0",
  "typescript": "~3.4.3"
}
~~~

To install the development dependency add **-dev** flag to installation string:

~~~
> npm i -dev new-package-name-to-install
~~~
