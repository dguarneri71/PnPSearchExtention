# PnP Modern Search - DG Extention

This project centralizes extensibility code made by my for the [PnP Modern Search solution (v4)](https://github.com/microsoft-search/pnp-modern-search). This project include:

- Custom web components
- Custom Handlebars helpers

This code is from [PnP Modern Search solution (v4)](https://github.com/microsoft-search/pnp-modern-search-extensibility-samples)

## Get Started

- Install and configure the [PnP Modern Search - Extensibility samples](https://microsoft-search.github.io/pnp-modern-search/installation/) in your SharePoint Online environment.

## Web Components

**Demo components:**  

- dg-custom-component
- dg-demo-component

**Generic components:**  

- msys-border
- msys-title-border
- msys-css-loader

## Example

```html
<msys-border 
    data-color="" 
    data-size="" 
    data-hide="" 
    data-class-name="" 
    data-css-url="">
    <template id="border-content">
        {content}
    </template>
</msys-border>
```  

| Property          | Description                                                                                                        |
| ----------------- | ------------------------------------------------------------------------------------------------------------------ |
| `data-color`      | The name for your component. This name will be used as the custom HTML element name (ex: `<my-custom-component>`). |
| `data-size`       | The web component class for that component.                                                                        |
| `data-hide`       | The web component class for that component.                                                                        |
| `data-class-name` | The web component class for that component.                                                                        |
| `data-css-url`    | The web component class for that component.                                                                        |
| `template`        | `<template id="border-content">`.                                                                                  |

```html
<msys-title-border 
    data-color="" 
    data-size="" 
    data-hide="" 
    data-hide-title="" 
    data-icon="" 
    data-title="" 
    data-title-bkg-color=""
    data-text-color="" 
    data-class-name="" 
    data-title-class-name="" 
    data-body-class-name="" 
    data-css-url="">
    <template id="border-content">
        {content}
    </template>
</msys-title-border>
```  

| Property          | Description                                                                                                        |
| ----------------- | ------------------------------------------------------------------------------------------------------------------ |
| `data-color`      | string - codice colore. |
| `data-size`       | string - number of px.                                                                        |
| `data-hide`       | boolean - true | false.                                                                        |
| `data-hide-title` | boolean - true | false.                                                                        |
| `data-icon`       | string - fabric ui icon name.                                                                        |
| `data-title`      | string.                                                                        |
| `data-title-bkg-color`      | string - codice colore.                                                                        |
| `data-text-color` | string - codice colore.                                                                        |
| `data-class-name` | string.                                                                        |
| `data-title-class-name` | string.                                                                        |
| `data-body-class-name` | string.                                                                        |
| `data-css-url`    | string.                                                                        |
| `template`        | `<template id="border-content"></template>`.                                                                                  |

```html
<msys-css-loader data-css-url=""></msys-css-loader>
```  

| Property          | Description                                                                                                        |
| ----------------- | ------------------------------------------------------------------------------------------------------------------ |
| `data-css-url`    | The web component class for that component.                                                                        |

Heandlebars Helpers:  

- cleanText
- getUrl
- getLocateUrl
- getDispUrl

### cleanText

Value is a string

```text
{{cleanText 'value'}}
```  

### getUrl

Mandatory properties: ServerRedirectedURL and Path  

Usage:  

```html
<a href="{{getUrl item}}">
```  

### getLocateUrl

Mandatory properties: ParentLink, SPWebUrl and ListItemID  

Usage:  

```html
<a href="{{getLocateUrl item}}">
```  

### getDispUrl

Mandatory properties: ParentLink, SPWebUrl and ListItemID  

Usage:  

```html
<a href="{{getDispUrl item}}">
```  

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
