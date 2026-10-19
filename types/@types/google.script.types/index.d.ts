declare namespace google {

    namespace script {

        export interface Runner {

            withSuccessHandler(handler: Function): Runner;

            withFailureHandler(handler: (error: Error) => void): Runner;

            withUserObject(object: any): Runner;

            ensureDate(value: any): void;

            fmtData(value: any, includeTime: boolean): void;

            normalizeTxt(text: any, cleanSeparators: boolean): void //string;

            extract<T extends Record<string, any>>(data: T[], columns: (string | keyof T)[]): void;

            merge<T extends Record<string, any>, U extends Record<string, any>>(target: T[], source: U[], keyTarget: string | keyof T, keySource?: string | keyof U): void;

        }

        export let run: Runner;

    }

}

