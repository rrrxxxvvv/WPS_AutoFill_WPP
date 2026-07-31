var WPS_Enum = {
  msoCTPDockPositionLeft: 0,
  msoCTPDockPositionRight: 2
}

function GetUrlPath() {
  let url = decodeURI(document.location.toString())
  if (url.indexOf('/') !== -1) {
    url = url.substring(0, url.lastIndexOf('/'))
  }
  return url
}
