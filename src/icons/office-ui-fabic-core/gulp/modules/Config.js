var path = require('path');
var pkg = require('../../package.json');
var Plugins = require('./Plugins');
var Utilities = require('./Utilities');

/**
 * Configuration class containing all properties to be used throughout the build
 */
var Config = function() {
  this.debugMode = false;
  this.sassExtension = "scss";
  this.buildSass = false;
  this.copyRightMessage = "Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.";
  
  this.paths = {
    dist: 'dist',
    src: 'src',
		componentsPath : 'src/components',
    srcLibPath: 'lib',
    temp: 'temp'
	};
  
  this.paths.distComponents = this.paths.dist + '/components';
  this.paths.distSass = this.paths.dist + '/sass';
  this.paths.distCSS = this.paths.dist + '/css';
  this.paths.distDocumentation = this.paths.dist + '/documentation';
  this.paths.distDocumentationCSS = this.paths.distDocumentation;
  this.paths.distSamples = this.paths.dist + '/samples';
  this.paths.distSampleComponents = this.paths.dist + '/samples/' + '/Components';
  this.paths.distJS = this.paths.dist + '/js';
  this.paths.distPackages = this.paths.dist + '/packages';
  this.paths.distDocumentation = this.paths.dist + '/documentation';
  this.paths.distDocsComponents = this.paths.distDocumentation + '/Components';
  this.paths.distDocsSamples = this.paths.distDocumentation + '/Samples';
  this.paths.distDocsStyles = this.paths.distDocumentation + '/Styles';
  
  this.paths.srcSamples = this.paths.src + '/samples';
  this.paths.srcData = this.paths.src + '/data';
  this.paths.srcSass = this.paths.src + '/sass';
  this.paths.srcDocs = this.paths.src + '/documentation';
  this.paths.srcDocsPages = this.paths.srcDocs + '/pages';
  this.paths.srcTemplate = this.paths.srcDocs + '/templates';
  this.paths.srcDocumentationCSS = this.paths.srcDocs + '/sass';
  this.paths.srcDocumentationModels = this.paths.srcTemplate + '/models';
  this.paths.srcDocTemplateModules = this.paths.srcTemplate + '/modules';
  this.paths.srcDocTemplateModulesComponents = this.paths.srcDocTemplateModules + '/components';
  
	this.port =  process.env.PORT || 2020;
	this.projectURL =  "http://localhost";
	this.projectDirectory =  path.resolve(__dirname, '../../' + this.paths.distDocumentation);
	this.servePaths = [
        {
            'urlPath': '/css',
            'folderPath': '../css'
        },
        {
            'urlPath': '/js',
            'folderPath': '../js'
        },
        {
            'urlPath': '/lib',
            'folderPath': '../lib'
        }
    ];
    this.typescriptConfig = {
        module: 'commonjs',
        declaration: true,
        target: 'ES5',
        noEmitOnError: true
    };
  this.typescriptProject = Plugins.tsc.createProject(this.typescriptConfig);
	this.nugetConfig = {
		id: "OfficeUIFabricCore",
		title: "Office UI Fabric Core",
		version: pkg.version,
		authors: "Microsoft",
		owners: "Microsoft",
		description: "Fabric is a responsive, mobile-first, front-end framework, designed to make it quick and simple for you to create web experiences using the Office Design Language. It’s easy to get up and running with Fabric—whether you’re creating a new Office experience from scratch or adding new features to an existing one.",
		summary: "The front-end framework for building experiences for Office and Office 365.",
		language: "en-us",
		projectUrl: "https://github.com/OfficeDev/Office-UI-Fabric-Core",
		licenseUrl: "https://github.com/OfficeDev/Office-UI-Fabric-Core/blob/master/LICENSE",
		copyright: "© Microsoft Corporation. All rights reserved.",
		requireLicenseAcceptance: true,
		tags: "Microsoft UI Fabric CSS",
		outputDir: this.paths.distPackages
	};
	this.nugetPaths = [
		{src: this.paths.distCSS, dest: "/content/Content/"},
		{src: this.paths.distSass, dest: "/content/Content/sass/"},
		{src: this.paths.distJS, dest: "/content/Scripts/"}
	];
  this.componentSamplesUpdate = "Components Samples updated successfully! Yay!";
  this.componentSamplesFinished = ' Component Samples build was successful! Yay!';
  //JS Files to be ignored in the JS Linter for Components
  //NOTE: Only use this for third party files, do not add any Fabric JS files to this.
  this.ignoreComponentJSLinting = [{
    src: this.paths.componentsPath + '/DatePicker/PickaDate.js',
    dist: this.paths.distSampleComponents + '/DatePicker/'
  }];
  //Errors
  this.genericBuildError = "Hmm, something went wrong in the build... Here is the error dump";
  this.templateLibraryConfig = {
    "components": [
      "ContextualHost",
      "Button"
    ]
  };
};

module.exports = new Config();
