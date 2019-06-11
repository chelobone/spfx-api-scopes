# Project Title

WebPart que permite consumir la API MSGraph para obtener documentos desde MSSharePoint y desplegarlos. Esta WebPart fue construida como prueba de concepto para usarla tanto en un contexto MSSharePoint como en MSTeams.

## Getting Started

See deployment for notes on how to deploy the project on a live system.

### Prerequisites

What things you need to install the software and how to install them

```
Give examples
```

### Installing

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

## Deployment

Para generar paquete de instalación en produccón

gulp build
gulp bundle --ship
gulp package-solution --ship

Se generará una carpeta en el directorio de la aplicación (sharepoint\solution), que contendrá un archivo de extensión .sppkg

## Built With

* [office-ui-fabric-react](https://developer.microsoft.com/en-us/fabric#/controls/web/) - Controles de office-ui-fabric-react

## Contributing


## Versioning



## Authors

* **Marcelo Friz** - *Initial work* - [chelobone](https://github.com/chelobone)

See also the list of [contributors](https://github.com/chelobone/spfx-api-scopes/contributors) who participated in this project.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Esto fue inspirado en la idea de tener un componente funcionando en contextos distintos





