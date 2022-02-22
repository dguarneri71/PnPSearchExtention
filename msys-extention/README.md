# PnP - Extensibility Library Demo Project

This project shows how to implement custom web component for the 'Search Results' Web Part.

## Documentation

- npm install @pnp/sp@2.0.3 --save //for SharePoint operations
- npm i xlsx //To parser and writer for spreadsheet formats (.xlsx file)
- npm i file-saver //For saving files on the client-side

## TODO

Gestire gli errori nel webcomponent Excel se non ci fosse la lista con le etichette indicata.

## Build

gulp clean; gulp bundle --ship; gulp package-solution --ship  
