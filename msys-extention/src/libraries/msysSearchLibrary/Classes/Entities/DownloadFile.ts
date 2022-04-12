export default class DownloadFile {
    public download: string;
    public filename: string;
    public webUrl: string; // current web url
    public serverRelativeUrl: string;
    public siteUrl: string; //file web url

    constructor(url: string, name: string, webUrl: string, siteUrl: string) {
        this.download = url;
        this.filename = name;
        this.webUrl = webUrl;
        this.siteUrl = siteUrl;
        this.serverRelativeUrl = "/" + (url.split("/")).slice(3).join("/");
    }
}