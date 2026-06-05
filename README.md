# Install the generator globally
npm install -g yo generator-office

# Run the generator
yo office

# Manifest validator
npx office-addin-manifest validate manifest.xml



# config
```
<!-- Your unique ID (keep yours) -->
<Id>7fae96f5-9626-4762-bfee-f42353b53783</Id>

<!-- Your branding -->
<ProviderName>SmartBlue</ProviderName>
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->

<DisplayName DefaultValue="Blue AI"/>
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->

<Description DefaultValue="AI Document Analysis for Outlook"/>
<SupportUrl DefaultValue="https://ws.demo.smartblue.ai"/>

<!-- Your GitHub Pages URLs -->
<IconUrl DefaultValue="https://lavsingh108.github.io/Outlook-Add-in/logo.png"/>
<HighResolutionIconUrl DefaultValue="https://lavsingh108.github.io/Outlook-Add-in/logo.png"/>

<!-- Your taskpane URL (appears multiple times — replace all) -->
<!-- Search for localhost:3000 and replace with: -->
https://lavsingh108.github.io/Outlook-Add-in/taskpane.html

<!-- Your app domains -->
<AppDomain>https://lavsingh108.github.io</AppDomain>
<AppDomain>https://ws.demo.smartblue.ai</AppDomain>
```