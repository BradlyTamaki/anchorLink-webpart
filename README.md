## anchor-link-webpart
SPFX SharePoint webpart to create anchor links on SharePoint modern page.

In a text webpart create a hyperlink with #anchorLink for address and it will jump to where the webpart is located on the page.

Note: it will not work outside the current page it lives on. 

![Screenshot](/assets/screenshot.png)

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
