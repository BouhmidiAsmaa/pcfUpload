import { IInputs, IOutputs } from "./generated/ManifestTypes";
import Dropzone from "dropzone";
import * as toastr from "toastr";

class AttachedFile implements ComponentFramework.FileObject {
    public fileId: string;
    public fileContent: string;
    public fileSize: number;
    public fileName: string;
    public mimeType: string;

    constructor(
        fileId: string,
        fileName: string,
        mimeType: string,
        fileContent: string,
        fileSize: number
    ) {
        this.fileId = fileId;
        this.fileName = fileName;
        this.mimeType = mimeType;
        this.fileContent = fileContent;
        this.fileSize = fileSize;
    }
}

interface FileItem {
    attachedFile: AttachedFile;
    dropzoneFile: any;
    container: HTMLDivElement;
}

export class PCFFileUploader
    implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _context!: ComponentFramework.Context<IInputs>;

    private _entityId = "";
    private _entityName = "";

    private _container!: HTMLDivElement;
    private _divDropZone!: HTMLDivElement;
    private _formDropZone!: HTMLFormElement;
    private _imgUpload!: HTMLImageElement;
    private _divFile!: HTMLDivElement;
    private _sendButton!: HTMLButtonElement;

    private _dropzone: Dropzone | null = null;

    private _attachedFiles: AttachedFile[] = [];
    private _fileItems: FileItem[] = [];

    private _urlSharepoint = "";

    private _userName = "";
    private _contactGuid = "";
    private _contactName = "";

    private _supportedfiles = "";
    private _environmentVariableFlow = "";

    constructor() {
        // Empty constructor
    }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {

        this._context = context;

        /*
         * Get current Dataverse record.
         */
        const page = (context as any).page;

        this._entityId = page?.entityId ?? "";
        this._entityName = page?.entityTypeName ?? "";

        /*
         * Remove braces from GUID.
         */
        this._entityId = this._entityId.replace(/[{}]/g, "");

        /*
         * Read PCF parameters.
         */
        this._supportedfiles =
            context.parameters.supporteFiles.raw ?? "";

        this._environmentVariableFlow =
            context.parameters.environmentVariableFlow.raw ?? "";

        /*
         * Current Dynamics user.
         */
        this._userName =
            Xrm.Utility.getGlobalContext().userSettings.userName ?? "";

        /*
         * Configure toastr.
         */
        toastr.options.closeButton = true;
        toastr.options.progressBar = true;
        toastr.options.positionClass = "toast-bottom-left";

        /*
         * Main PCF container.
         */
        this._container = document.createElement("div");
        this._container.className = "pcf-file-uploader";

        container.appendChild(this._container);

        /*
         * Create controls.
         */
        this.createDropzoneUI();
        this.createFileListUI();
        this.createSendButton();

        /*
         * Load Power Automate URL first.
         */
        void this.initializeControl();
    }

    /**
     * Initialize Flow URL then Dropzone.
     */
    private async initializeControl(): Promise<void> {

        try {

            await this.setFlowURL(
                this._environmentVariableFlow
            );

            this.initializeDropzone();

        } catch (error) {

            console.error(
                "Unable to initialize uploader:",
                error
            );

            toastr.error(
                "The upload service could not be initialized."
            );
        }
    }

    /**
     * Create Dropzone UI.
     */
    private createDropzoneUI(): void {

        this._divDropZone =
            document.createElement("div");

        this._divDropZone.id = "dropzone";

        this._formDropZone =
            document.createElement("form");

        this._formDropZone.id = "upload_dropzone";
        this._formDropZone.method = "post";
        this._formDropZone.className =
            "dropzone needsclick";

        /*
         * Prevent browser form submission.
         */
        this._formDropZone.addEventListener(
            "submit",
            (event: Event) => {
                event.preventDefault();
            }
        );

        /*
         * Upload image.
         */
        this._imgUpload =
            document.createElement("img");

        this._imgUpload.className =
            "uploadImgHowerOut";

        this._imgUpload.alt =
            "Upload files";

        this._context.resources.getResource(
            "newUploadIcon.png",

            (data: string) => {

                this._imgUpload.src =
                    this.generateSrcUrl(
                        "image",
                        "png",
                        data
                    );
            },

            () => {

                console.error(
                    "Unable to load upload image."
                );
            }
        );

        /*
         * Hover effect.
         */
        this._imgUpload.addEventListener(
            "mouseover",
            () => {

                this._imgUpload.className =
                    "uploadImgHower";
            }
        );

        this._imgUpload.addEventListener(
            "mouseout",
            () => {

                this._imgUpload.className =
                    "uploadImgHowerOut";
            }
        );

        /*
         * Dropzone message.
         */
        const message =
            document.createElement("div");

        message.className = "dz-message";

        message.textContent =
            "Drag files here or click to select files";

        this._formDropZone.appendChild(
            this._imgUpload
        );

        this._formDropZone.appendChild(
            message
        );

        this._divDropZone.appendChild(
            this._formDropZone
        );

        this._container.appendChild(
            this._divDropZone
        );
    }

    /**
     * File list container.
     */
    private createFileListUI(): void {

        this._divFile =
            document.createElement("div");

        this._divFile.className =
            "selected-files";

        this._container.appendChild(
            this._divFile
        );
    }

    /**
     * Upload button.
     */
    private createSendButton(): void {

        this._sendButton =
            document.createElement("button");

        this._sendButton.type = "button";

        this._sendButton.innerText =
            "Upload to SharePoint";

        this._sendButton.className =
            "boton";

        this._sendButton.style.marginTop =
            "10px";

        this._sendButton.style.display =
            "none";

        this._sendButton.addEventListener(
            "click",
            () => {

                void this.handleUpload();
            }
        );

        this._container.appendChild(
            this._sendButton
        );
    }

    /**
     * Generate image URL.
     */
    private generateSrcUrl(
        datatype: string,
        fileType: string,
        fileContent: string
    ): string {

        return `data:${datatype}/${fileType};base64,${fileContent}`;
    }

    /**
     * Initialize Dropzone.
     */
    private initializeDropzone(): void {

        /*
         * Disable Dropzone automatic discovery.
         */
        Dropzone.autoDiscover = false;

        /*
         * Dropzone expects string | undefined,
         * NOT string | null.
         */
        const acceptedFiles: string | undefined =
            this._supportedfiles.trim().length > 0
                ? this._supportedfiles
                : undefined;

        this._dropzone =
            new Dropzone(
                this._formDropZone,
                {
                    /*
                     * Required by Dropzone even though
                     * autoProcessQueue is disabled.
                     */
                    url: "/",

                    acceptedFiles: acceptedFiles,

                    /*
                     * Maximum file size in MB.
                     */
                    maxFilesize: 3,

                    /*
                     * Maximum number of selected files.
                     */
                    maxFiles: 20,

                    parallelUploads: 2,

                    filesizeBase: 1000,

                    /*
                     * We upload manually to Flow.
                     */
                    autoProcessQueue: false,

                    addRemoveLinks: false,

                    clickable: true,

                    /*
                     * Don't let Dropzone create its
                     * default preview UI.
                     *
                     * This is compatible with the
                     * Dropzone typings in your project.
                     */
                    previewsContainer:
                        false as unknown as string,

                    dictDefaultMessage:
                        "Drag files here or click to select files",

                    dictInvalidFileType:
                        "Unsupported file type.",

                    dictFileTooBig:
                        "File is too large. Maximum size is 3 MB.",

                    dictMaxFilesExceeded:
                        "Maximum 20 files are allowed."
                }
            );

        /*
         * File selected.
         */
        this._dropzone.on(
            "addedfile",
            (file: any) => {

                /*
                 * Allow Dropzone validation to complete
                 * before reading the file.
                 */
                setTimeout(
                    () => {

                        if (file.accepted === false) {
                            return;
                        }

                        this.readFile(file);
                    },
                    0
                );
            }
        );

        /*
         * Validation error.
         */
        this._dropzone.on(
            "error",
            (
                file: any,
                message: string | object
            ) => {

                const errorMessage =
                    typeof message === "string"
                        ? message
                        : "The selected file is not valid.";

                toastr.error(errorMessage);

                this.removeDropzoneFile(file);
            }
        );
    }

    /**
     * Convert selected file to Base64.
     */
    private readFile(file: any): void {

        /*
         * Prevent duplicate processing.
         */
        const alreadyExists =
            this._fileItems.some(
                item =>
                    item.dropzoneFile === file
            );

        if (alreadyExists) {
            return;
        }

        const reader =
            new FileReader();

        reader.onload = () => {

            try {

                const result =
                    reader.result;

                if (
                    typeof result !== "string"
                ) {

                    throw new Error(
                        "Unable to convert file to Base64."
                    );
                }

                const base64Marker =
                    ";base64,";

                const index =
                    result.indexOf(
                        base64Marker
                    );

                if (index === -1) {

                    throw new Error(
                        "Invalid Base64 content."
                    );
                }

                /*
                 * Remove:
                 *
                 * data:application/pdf;base64,
                 *
                 * Keep only actual Base64.
                 */
                const fileContent =
                    result.substring(
                        index +
                        base64Marker.length
                    );

                const attachedFile =
                    new AttachedFile(
                        "",
                        file.name,
                        file.type ||
                        "application/octet-stream",
                        fileContent,
                        file.size
                    );

                this._attachedFiles.push(
                    attachedFile
                );

                this.addFileControl(
                    attachedFile,
                    file
                );

                this.toggleSendButtonVisibility();

            } catch (error) {

                console.error(
                    "Unable to read file:",
                    error
                );

                toastr.error(
                    `Unable to read ${file.name}.`
                );

                this.removeDropzoneFile(
                    file
                );
            }
        };

        reader.onerror = () => {

            console.error(
                `FileReader failed for ${file.name}.`
            );

            toastr.error(
                `Unable to read ${file.name}.`
            );

            this.removeDropzoneFile(
                file
            );
        };

        reader.readAsDataURL(file);
    }

    /**
     * Display selected file.
     */
    private addFileControl(
        file: AttachedFile,
        dropzoneFile: any
    ): void {

        const fileContainer =
            document.createElement("div");

        fileContainer.className =
            "fileContainer";

        fileContainer.style.display =
            "flex";

        fileContainer.style.alignItems =
            "center";

        fileContainer.style.justifyContent =
            "space-between";

        fileContainer.style.marginBottom =
            "5px";

        /*
         * File name.
         */
        const fileLabel =
            document.createElement("span");

        fileLabel.innerText =
            file.fileName;

        fileLabel.className =
            "text-font";

        fileLabel.style.marginRight =
            "10px";

        /*
         * File size.
         */
        const sizeLabel =
            document.createElement("span");

        sizeLabel.innerText =
            this.formatFileSize(
                file.fileSize
            );

        sizeLabel.style.marginLeft =
            "auto";

        sizeLabel.style.marginRight =
            "15px";

        /*
         * Cancel button.
         */
        const removeButton =
            document.createElement("button");

        removeButton.type =
            "button";

        removeButton.innerText =
            "Cancel";

        removeButton.className =
            "boton";

        removeButton.addEventListener(
            "click",
            () => {

                this.removeSelectedFile(
                    file,
                    dropzoneFile,
                    fileContainer
                );
            }
        );

        fileContainer.appendChild(
            fileLabel
        );

        fileContainer.appendChild(
            sizeLabel
        );

        fileContainer.appendChild(
            removeButton
        );

        this._divFile.appendChild(
            fileContainer
        );

        /*
         * Keep relationship between our
         * attachment and Dropzone file.
         */
        this._fileItems.push({
            attachedFile: file,
            dropzoneFile: dropzoneFile,
            container: fileContainer
        });
    }

    /**
     * Remove selected file.
     */
    private removeSelectedFile(
        attachedFile: AttachedFile,
        dropzoneFile: any,
        fileContainer: HTMLDivElement
    ): void {

        /*
         * Remove from our upload array.
         */
        const attachmentIndex =
            this._attachedFiles.indexOf(
                attachedFile
            );

        if (
            attachmentIndex !== -1
        ) {

            this._attachedFiles.splice(
                attachmentIndex,
                1
            );
        }

        /*
         * Remove FileItem mapping.
         */
        const itemIndex =
            this._fileItems.findIndex(
                item =>
                    item.attachedFile ===
                    attachedFile
            );

        if (
            itemIndex !== -1
        ) {

            this._fileItems.splice(
                itemIndex,
                1
            );
        }

        /*
         * Remove from Dropzone itself.
         *
         * Important because maxFiles = 20.
         */
        this.removeDropzoneFile(
            dropzoneFile
        );

        /*
         * Remove visual item.
         */
        if (
            fileContainer.parentElement
        ) {

            fileContainer.parentElement
                .removeChild(
                    fileContainer
                );
        }

        this.toggleSendButtonVisibility();
    }

    /**
     * Remove file safely from Dropzone.
     */
    private removeDropzoneFile(
        file: any
    ): void {

        if (
            !this._dropzone ||
            !file
        ) {
            return;
        }

        try {

            this._dropzone.removeFile(
                file
            );

        } catch (error) {

            console.warn(
                "Unable to remove file from Dropzone:",
                error
            );
        }
    }

    /**
     * User clicked Upload.
     */
    private async handleUpload(): Promise<void> {

        if (
            this._attachedFiles.length === 0
        ) {

            toastr.warning(
                "Please select at least one file."
            );

            return;
        }

        if (
            !this._urlSharepoint
        ) {

            toastr.error(
                "Upload service is not available."
            );

            return;
        }

        /*
         * Take a snapshot of selected files.
         */
        const filesToUpload =
            [...this._attachedFiles];

        this.setUploadingState(
            true
        );

        try {

            await this.uploadFileToSP(
                filesToUpload,
                ""
            );

            /*
             * Only show success AFTER Flow
             * actually returned success.
             */
            toastr.success(
                `${filesToUpload.length} file(s) uploaded successfully.`
            );

            /*
             * Only clear files after successful
             * Power Automate response.
             */
            this.clearFiles();

        } catch (error) {

            console.error(
                "SharePoint upload failed:",
                error
            );

            const message =
                error instanceof Error
                    ? error.message
                    : "Upload failed.";

            toastr.error(
                message
            );

        } finally {

            this.setUploadingState(
                false
            );
        }
    }

    /**
     * Send files to Power Automate.
     */
    private uploadFileToSP(
        files: AttachedFile[],
        noteText: string
    ): Promise<void> {

        return new Promise<void>(
            (resolve, reject) => {

                if (
                    !this._urlSharepoint
                ) {

                    reject(
                        new Error(
                            "Power Automate URL is empty."
                        )
                    );

                    return;
                }

                const req =
                    new XMLHttpRequest();

                req.open(
                    "POST",
                    this._urlSharepoint,
                    true
                );

                req.setRequestHeader(
                    "Content-Type",
                    "application/json"
                );

                /*
                 * 2 minute timeout.
                 */
                req.timeout =
                    120000;

                /*
                 * Flow response.
                 */
                req.onload = () => {

                    if (
                        req.status >= 200 &&
                        req.status < 300
                    ) {

                        resolve();

                        return;
                    }

                    console.error(
                        "Flow returned an error:",
                        {
                            status:
                                req.status,

                            response:
                                req.responseText
                        }
                    );

                    reject(
                        new Error(
                            `Upload failed. HTTP status ${req.status}.`
                        )
                    );
                };

                /*
                 * Network failure.
                 */
                req.onerror = () => {

                    reject(
                        new Error(
                            "Network error while calling Power Automate."
                        )
                    );
                };

                /*
                 * Timeout.
                 */
                req.ontimeout = () => {

                    reject(
                        new Error(
                            "The upload request timed out."
                        )
                    );
                };

                /*
                 * Payload sent to Power Automate.
                 */
                const payload = {

                    files:
                        files.map(
                            file => ({
                                filename:
                                    file.fileName,

                                filesize:
                                    file.fileSize.toString(),

                                mimetype:
                                    file.mimeType,

                                documentbody:
                                    file.fileContent
                            })
                        ),

                    contactId:
                        this._contactGuid || "",

                    userName:
                        this._userName || "",

                    noteText:
                        noteText || "",

                    recordId:
                        this._entityId || "",

                    contactName:
                        this._contactName || ""
                };

                console.log(
                    `Uploading ${files.length} file(s) for record ${this._entityId}`
                );

                req.send(
                    JSON.stringify(
                        payload
                    )
                );
            }
        );
    }

    /**
     * Get Power Automate URL from
     * Dataverse environment variable.
     */
    private async setFlowURL(
        environmentVariableFlow: string
    ): Promise<void> {

        if (
            !environmentVariableFlow
        ) {

            throw new Error(
                "Environment variable schema name is empty."
            );
        }

        /*
         * Escape value for FetchXML.
         */
        const safeSchemaName =
            environmentVariableFlow
                .replace(
                    /&/g,
                    "&amp;"
                )
                .replace(
                    /</g,
                    "&lt;"
                )
                .replace(
                    />/g,
                    "&gt;"
                )
                .replace(
                    /"/g,
                    "&quot;"
                )
                .replace(
                    /'/g,
                    "&apos;"
                );

        const fetchXml = `
            <fetch
                version="1.0"
                output-format="xml-platform"
                mapping="logical"
                distinct="false"
                top="1">

                <entity name="environmentvariablevalue">

                    <attribute name="environmentvariablevalueid" />
                    <attribute name="value" />
                    <attribute name="createdon" />

                    <order
                        attribute="createdon"
                        descending="true" />

                    <link-entity
                        name="environmentvariabledefinition"
                        from="environmentvariabledefinitionid"
                        to="environmentvariabledefinitionid"
                        link-type="inner"
                        alias="definition">

                        <filter type="and">

                            <condition
                                attribute="schemaname"
                                operator="eq"
                                value="${safeSchemaName}" />

                        </filter>

                    </link-entity>

                </entity>

            </fetch>
        `;

        try {

            const result =
                await this._context.webAPI
                    .retrieveMultipleRecords(
                        "environmentvariablevalue",
                        `?fetchXml=${encodeURIComponent(fetchXml)}`
                    );

            if (
                result.entities.length === 0
            ) {

                throw new Error(
                    `Environment variable '${environmentVariableFlow}' has no current value.`
                );
            }

            const flowUrl =
                result.entities[0].value;

            if (
                typeof flowUrl !== "string" ||
                flowUrl.trim().length === 0
            ) {

                throw new Error(
                    `Environment variable '${environmentVariableFlow}' contains an empty Flow URL.`
                );
            }

            this._urlSharepoint =
                flowUrl.trim();

        } catch (error) {

            console.error(
                "Error retrieving environment variable:",
                error
            );

            throw error;
        }
    }

    /**
     * Clear all selected files.
     */
    private clearFiles(): void {

        this._attachedFiles = [];
        this._fileItems = [];

        this._divFile.innerHTML =
            "";

        if (
            this._dropzone
        ) {

            try {

                this._dropzone.removeAllFiles(
                    true
                );

            } catch (error) {

                console.warn(
                    "Unable to clear Dropzone:",
                    error
                );
            }
        }

        this.toggleSendButtonVisibility();
    }

    /**
     * Show Upload button only when
     * files have been selected.
     */
    private toggleSendButtonVisibility(): void {

        this._sendButton.style.display =
            this._attachedFiles.length > 0
                ? "inline-block"
                : "none";
    }

    /**
     * Upload state.
     */
    private setUploadingState(
        uploading: boolean
    ): void {

        this._sendButton.disabled =
            uploading;

        this._sendButton.innerText =
            uploading
                ? "Uploading..."
                : "Upload to SharePoint";

        /*
         * Prevent file changes while uploading.
         */
        if (
            this._dropzone
        ) {

            if (
                uploading
            ) {

                this._dropzone.disable();

            } else {

                this._dropzone.enable();
            }
        }
    }

    /**
     * Format file size.
     */
    private formatFileSize(
        bytes: number
    ): string {

        if (
            bytes < 1000
        ) {

            return `${bytes} B`;
        }

        if (
            bytes < 1000000
        ) {

            return `${(
                bytes / 1000
            ).toFixed(1)} KB`;
        }

        return `${(
            bytes / 1000000
        ).toFixed(1)} MB`;
    }

    /**
     * PCF updateView.
     */
    public updateView(
        context: ComponentFramework.Context<IInputs>
    ): void {

        this._context =
            context;
    }

    /**
     * PCF outputs.
     */
    public getOutputs(): IOutputs {

        return {};
    }

    /**
     * PCF cleanup.
     */
    public destroy(): void {

        if (
            this._dropzone
        ) {

            try {

                this._dropzone.destroy();

            } catch (error) {

                console.warn(
                    "Unable to destroy Dropzone:",
                    error
                );
            }

            this._dropzone =
                null;
        }
    }
}