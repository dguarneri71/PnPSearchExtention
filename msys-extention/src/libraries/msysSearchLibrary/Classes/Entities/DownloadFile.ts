export default class DownloadFile {
    public download: string;
    public filename: string;
    public webUrl: string;

    constructor(url: string, name: string, webUrl: string) {
        this.download = url;
        this.filename = name;
        this.webUrl = webUrl;
    }
}