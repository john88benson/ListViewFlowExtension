## flow-pompt-cmd-set

### Debug URL Content Curation
?loadSpfx=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"4db1937b-cb36-44ba-9ef9-3346ae9d6be2":{"location":"ClientSideExtension.ListViewCommandSet.ContextMenu","properties":{"sampleTextOne":"One item is selected in the list."}}}


### Build Deployment Package Content Curation
gulp clean
gulp bundle --ship
gulp package-solution --ship
