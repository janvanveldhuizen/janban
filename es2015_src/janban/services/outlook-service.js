class OutlookService {
    constructor() {
    }

    isRunningInOutlook() {
        alert('function isRunningInOutlook')
        return (window.external !== undefined && window.external.OutlookApplication !== undefined);
    }
}

// export default OutlookService;