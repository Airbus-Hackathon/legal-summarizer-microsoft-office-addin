# legal-summarizer-microsoft-office-addin

A microsoft word office addin to perform legal document summarization from withing microsoft word application.

In order to run the add in you need a microsoft office activated license.

This addin acts as a static web application and must be made accessible over internet or in a private network

## Architecture

[Global Architecture](./doc/global_architecture.png)
[Sequence Diagram](./doc/sequence_diagram.png)

## Resources

[Manifest file](./doc/manifest.xml)

## Build

```bash
# build the code
npm run build
# compress the content
tar -czf ./archive/legal-summarizer-microsoft-office-addin.tar.gz ./dist
# replace host in manifest.xml according to your production settings
# https://www.contoso.com => https://legal-summarizer.students-epitech.ovh
```

## Run in development

```bash
# launch the local development server on port 3000
npm run start
# kill the local development server
lsof -i tcp:3000
kill -9 <PID>
```

## Usage

[Use the Add-In Video](./doc/tutorials/add-in-usage.mov)

## Setup in production

[Manually sideload the addin official documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#sideload-a-yeoman-created-add-in-to-office-on-the-web)

[Setup the Add-In Video](./doc/tutorials/add-in-setup.mov)

![Microsoft Add-in Setup 1/15](./doc/add-in-setup/setup-1-15.png?raw=true "Microsoft Add-in Setup 1/15")
![Microsoft Add-in Setup 2/15](./doc/add-in-setup/setup-2-15.png?raw=true "Microsoft Add-in Setup 2/15")
![Microsoft Add-in Setup 3/15](./doc/add-in-setup/setup-3-15.png?raw=true "Microsoft Add-in Setup 3/15")
![Microsoft Add-in Setup 4/15](./doc/add-in-setup/setup-4-15.png?raw=true "Microsoft Add-in Setup 4/15")
![Microsoft Add-in Setup 5/15](./doc/add-in-setup/setup-5-15.png?raw=true "Microsoft Add-in Setup 5/15")
![Microsoft Add-in Setup 6/15](./doc/add-in-setup/setup-6-15.png?raw=true "Microsoft Add-in Setup 6/15")
![Microsoft Add-in Setup 7/15](./doc/add-in-setup/setup-7-15.png?raw=true "Microsoft Add-in Setup 7/15")
![Microsoft Add-in Setup 8/15](./doc/add-in-setup/setup-8-15.png?raw=true "Microsoft Add-in Setup 8/15")
![Microsoft Add-in Setup 9/15](./doc/add-in-setup/setup-9-15.png?raw=true "Microsoft Add-in Setup 9/15")
![Microsoft Add-in Setup 10/15](./doc/add-in-setup/setup-10-15.png?raw=true "Microsoft Add-in Setup 10/15")
![Microsoft Add-in Setup 11/15](./doc/add-in-setup/setup-11-15.png?raw=true "Microsoft Add-in Setup 11/15")
![Microsoft Add-in Setup 12/15](./doc/add-in-setup/setup-12-15.png?raw=true "Microsoft Add-in Setup 12/15")
![Microsoft Add-in Setup 13/15](./doc/add-in-setup/setup-13-15.png?raw=true "Microsoft Add-in Setup 13/15")
![Microsoft Add-in Setup 14/15](./doc/add-in-setup/setup-14-15.png?raw=true "Microsoft Add-in Setup 14/15")
![Microsoft Add-in Setup 15/15](./doc/add-in-setup/setup-15-15.png?raw=true "Microsoft Add-in Setup 15/15")

## Docker images 

[legal-summarizer API](https://hub.docker.com/r/antoineleguillou/legal-summarizer)

## Dependencies

[Microsoft Word AddIn Static files](https://github.com/Airbus-Hackathon/legal-summarizer-microsoft-office-addin/tree/main/archive)
[legal-summarizer API](https://github.com/Airbus-Hackathon/ai-api-bart-fintuned)
