import {siteCollectionUrl, cdnAssetsBaseUrl} from "./SharePoint"

export function getQueryParameters(url = window.location.search): any {
  if (!url) return {}
  return (
    url.substring(url.indexOf("?") + 1)
      .split("&")
      .reduce((params, kv) => ({
        ...params,
        [kv.split("=")[0]]: decodeURIComponent(kv.split("=")[1])
      }), {})
  )
}

export function createQueryParameters(obj = {}) {
  return Object.keys(obj).reduce((query, key, i) => (
    query +
    (i === 0 ? "?" : "&") +
    `${key}=${encodeURIComponent(obj[key])}`
  ), "")
}

// This came from a StackOverflow answer
export function humanFileSize(size: number) {
  var i = size === 0 ? 0 : Math.floor(Math.log(size) / Math.log(1024))
  return parseFloat((size / Math.pow(1024, i)).toFixed(2)) * 1 + ' ' + ['B', 'KB', 'MB', 'GB', 'TB'][i]
}

const officeUiFabricSupportedFileIcons = {
  accdb: true,
  csv: true,
  docx: true,
  dotx: true,
  mpp: true,
  mpt: true,
  odp: true,
  ods: true,
  odt: true,
  one: true,
  onepkg: true,
  onetoc: true,
  potx: true,
  ppsx: true,
  pptx: true,
  pub: true,
  vsdx: true,
  vssx: true,
  vstx: true,
  xls: true,
  xlsx: true,
  xltx: true,
  xsn: true
}

export function getGenericIconPath(extension: string) {
  let path = cdnAssetsBaseUrl + "/images/"
  
  switch (extension.toLowerCase()) {
    case "zip": 
    case "rar":
    case "7z":
    case "iso":
    case "cab":
        path += "zip.png"
      break
    case "mp3":
    case "aac":
    case "wav":
    case "aiff":
    case "ogg":
    case "wma":
    case "flac":
    case "alac":
    case "aax":
    case "m4a":
    case "m4b":
      path += "audio.png"
      break
    case "":
      path += "blank.png"
      break
    case "php":
    case "js":
    case "go":
    case "cpp":
    case "c":
    case "json":
    case "sql":
    case "csharp":
    case "bat":
    case "fsharp":
    case "less":
    case "sass":
    case "ruby":
    case "perl":
    case "perl6":
    case "powershell":
    case "ps":
    case "scss":
    case "ts":
    case "typescript":
    case "xsl":
    case "markdown":
      path += "code.png"
      break
    case "csv":
      path += "csv.png"
      break
    case "exe":
      path += "exe.png"
      break
    case "html":
      path += "html.png"
      break
    case "jpg":
    case "png":
    case "jpeg":
    case "gif":
    case "tiff":
    case "bmp":
    case "svg":
      path += "photo.png"
      break
    case "avi":
    case "mpeg":
    case "mkv":
    case "mp4":
    case "mov":
    case "qt":
    case "wmv":
    case "webm":
    case "m4p":
    case "mpg":
    case "mp2":
    case "mpv":
    case "m4v":
    case "3gp":
    case "3g2":
      path += "video.png"
      break
    default: 
      path += "genericfile.png"
      break
  }
  return path;
}

const officeUiFabricIconUrl = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/{{ext}}_16x1.svg"

export function getFileExt(filename) {
  const i = filename.lastIndexOf(".")
  return i >= 0
    ? filename.substring(i + 1).toLowerCase()
    : ""
}

export function getOfficeUiFabricFileIconUrl(filename: string): string {
  const ext = getFileExt(filename)
  return officeUiFabricSupportedFileIcons[ext]
    ? officeUiFabricIconUrl.replace("{{ext}}", ext)
    : getGenericIconPath(ext)
}

export function sortBy(collection, prop, order = "asc") {
  collection.sort((a, b) => {
    const x = a[prop]
    const y = b[prop]
    return order === "asc"
      ? (x > y ? 1 : x < y ? -1 : 0)
      : (x < y ? 1 : x > y ? -1 : 0)
  })
}

export function sortMomentsBy(collection, prop, order = "asc") {
  collection.sort((a, b) => {
    const x = a[prop]
    const y = b[prop]
    return order === "asc"
      ? x.diff(y)
      : y.diff(x)
  })
}

export function fixPageStyling() {
  // This fix will
  // - hide the default page title/hero, spacer, and comments section
  // - expand the web part width to full instead of the default max width of 1280px
  // - remove the page's bottom margin and padding
  // - disable vertical page scrolling
  //const style = document.createElement("style")
 // style.innerHTML = (`
  /*div[class^="pageTitle"],
    div[class^="canvasSpacerSection"],
    div[class^="commentsWrapper"] {
      display: none;
    }
    .CanvasZone {
      max-width: initial;
    }
    div.CanvasSection > div.ControlZone {
      margin-bottom: 0 !important;
      padding-bottom: 0 !important;
    }
    
    div[class^="scrollRegion"] {
      overflow-y: hidden;
    }
    */
  //`)
  //document.head.appendChild(style)
}
