import {
    DataMashup,
    MashupBinaryPrefix,
    MashupBinaryRegExp,
    MashupFormulaSectionDefault,
} from './datamashup';
import type { Result } from './types';
import { type Uint8ArrayUtilsFrom, Uint8ArrayUtils } from './utils';
import { type UnzippedItem, Unzip, Zip } from './zip';

export class ExcelCustomXml {
    private _xmlContent: string;
    private _datamashup: DataMashup;

    private get mashupBase64(): string | undefined {
        const match = this._xmlContent.match(MashupBinaryRegExp);
        return match ? match[1] : undefined;
    }

    private set mashupBase64(value: string) {
        if (!value) {
            return;
        }
        const current = this.mashupBase64;
        if (!current || current === value) {
            return;
        }
        this._xmlContent = this._xmlContent.replace(current, value);
    }

    public get datamashup(): DataMashup {
        return this._datamashup;
    }

    private constructor(xmlContent: string) {
        this._xmlContent = xmlContent;
        this._datamashup = undefined as never; // the static `create` ensures this is defined
    }

    private async unpack(): Promise<void> {
        const base64 = this.mashupBase64;
        const array = base64
            ? Uint8ArrayUtils.fromBase64(base64)
            : new Uint8Array();
        this._datamashup = await DataMashup.unpack(array);
    }

    public async pack(): Promise<string | undefined> {
        if (!this._datamashup) {
            return;
        }
        const result = await this._datamashup.pack();
        if (!result) {
            return;
        }
        this.mashupBase64 = Uint8ArrayUtils.toBase64(result);
        return this._xmlContent;
    }

    public static async create(xmlContent: string): Promise<ExcelCustomXml> {
        const instance = new this(xmlContent);
        await instance.unpack();
        return instance;
    }
}

export class ZipFile {
    private readonly zipData: Uint8Array;
    private _zipItems: UnzippedItem[] | undefined;

    public get zipItems(): UnzippedItem[] | undefined {
        return this._zipItems;
    }

    constructor(zipData: Uint8Array) {
        this.zipData = zipData;
    }

    public async unzip(): Promise<Result<UnzippedItem[]>> {
        if (this._zipItems) {
            return {
                ok: true,
                data: this._zipItems,
            };
        }
        const data = Uint8ArrayUtils.toNumberArray(this.zipData);
        const zipResult = await Unzip(data);
        if (!zipResult.ok) {
            return zipResult;
        }
        this._zipItems = zipResult.data;
        return {
            ok: true,
            data: this._zipItems,
        };
    }

    public async zip(): Promise<Result<Uint8Array>> {
        if (!this._zipItems) {
            return {
                ok: true,
                data: this.zipData,
            };
        }
        const items = this._zipItems.map((item) => {
            if (!item.encoding) {
                return item;
            }
            const data = Uint8ArrayUtils.fromStringEncoding(
                item.data,
                item.encoding,
                item.bom
            );
            return { ...item, data };
        });
        const zipResult = await Zip(items);
        if (!zipResult.ok) {
            return zipResult;
        }
        return {
            ok: true,
            data: zipResult.data,
        };
    }

    public setFileContents(
        item: UnzippedItem,
        data: UnzippedItem['data']
    ): Result<void> {
        const zipItem = this.zipItems && this.zipItems.find((o) => o === item);
        if (!zipItem) {
            return {
                ok: false,
                error: 'File not found in zip.',
            };
        }
        zipItem.data = data;
        return {
            ok: true,
            data: undefined,
        };
    }
}

export class ExcelZip extends ZipFile {
    private _mashupItem: UnzippedItem | null | undefined;
    private _mashup: Promise<ExcelCustomXml> | undefined;
    private _powerQueryItems: UnzippedItem[] | undefined;

    public get zipItems(): UnzippedItem[] {
        return super.zipItems || [];
    }

    private get mashupItem(): UnzippedItem | undefined {
        if (this._mashupItem !== undefined) {
            return this._mashupItem || undefined;
        }
        this._mashupItem = this.getMashupFile();
        return this._mashupItem;
    }

    public get mashup(): Promise<ExcelCustomXml> | undefined {
        if (this._mashup) {
            return this._mashup;
        }
        if (!this.mashupItem) {
            return undefined;
        }
        const data = this.mashupItem.data as string;
        this._mashup = ExcelCustomXml.create(data);
        return this._mashup;
    }

    constructor(zipData: Uint8ArrayUtilsFrom) {
        zipData = Uint8ArrayUtils.from(zipData);
        super(zipData);
    }

    private convertItemToString(item: UnzippedItem): UnzippedItem {
        if (item.type !== 'File') {
            return item;
        }
        const { path } = item;
        const isXml = path.endsWith('.xml') || path.endsWith('.m');
        if (!isXml) {
            return item;
        }
        if (typeof item.data === 'string') {
            return item;
        }
        const [data, encoding, bom] = Uint8ArrayUtils.toStringEncoding(
            item.data
        );
        return { ...item, data, encoding, bom };
    }

    public async unzip(): Promise<Result<UnzippedItem[]>> {
        const result = await super.unzip();
        if (!result.ok) {
            return result;
        }
        const { data } = result;
        for (let i = 0; i < data.length; i++) {
            let item = data[i];
            item = this.convertItemToString(item);
            data[i] = item;
        }
        await this.mashup;
        return {
            ok: true,
            data,
        };
    }

    public async zip(): Promise<Result<Uint8Array>> {
        const mashupItem = this.mashupItem;
        const mashup = await this.mashup;
        if (!mashupItem || !mashup) {
            return super.zip();
        }
        mashup.datamashup.resetPermissions();
        const xml = await mashup.pack();
        if (!xml) {
            return {
                ok: false,
                error: 'Unable to serialize CustomXml MashupData object.',
            };
        }
        const fileResult = this.setFileContents(mashupItem, xml);
        if (!fileResult.ok) {
            return fileResult;
        }
        return super.zip();
    }

    private getMashupFile(): UnzippedItem | undefined {
        return this.zipItems.find(({ type, path, data }) => {
            if (
                type !== 'File' ||
                typeof data !== 'string' ||
                !path.includes('customXml') ||
                !path.includes('item')
            ) {
                return false;
            }
            return data.includes(MashupBinaryPrefix);
        });
    }

    public async getPowerQueryFiles(): Promise<UnzippedItem[] | undefined> {
        if (this._powerQueryItems) {
            return this._powerQueryItems;
        }
        const mashup = await this.mashup;
        if (!mashup) {
            return;
        }
        const items = mashup.datamashup.rootItems;
        if (!items) {
            return;
        }
        this._powerQueryItems = items.filter((o) => o.path.endsWith('.m'));
        return this._powerQueryItems;
    }

    public async getPowerQueryFile(): Promise<UnzippedItem | undefined> {
        const items = await this.getPowerQueryFiles();
        if (!items) {
            return;
        }
        const item = items.find((o) =>
            o.path.endsWith(MashupFormulaSectionDefault)
        );
        return item;
    }

    public async setPowerQueryFile(
        item: UnzippedItem,
        data: UnzippedItem['data']
    ): Promise<void> {
        const mashup = await this.mashup;
        if (!mashup) {
            return;
        }
        mashup.datamashup.setFileContents(item, data);
    }

    public static async unzip(zipData: Uint8ArrayUtilsFrom): Promise<ExcelZip> {
        const instance = new this(zipData);
        await instance.unzip();
        return instance;
    }
}
