declare module 'deep-rename-keys' {
    function deepRenameKeys(obj: any, cb: (key: string) => string): any;

    export = deepRenameKeys;
}