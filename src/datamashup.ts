import { Parser } from 'binary-parser';

/**
 * This struct matches the top-level binary stream.
 *
 * References:
 *
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/22557f6d-7c29-4554-8fe4-7b7a54ac7a2b
 */
export const ParserRoot = new Parser()
    .endianness('little')
    .uint32le('version')
    .uint32le('packagePartsLength')
    .array('packageParts', { type: 'uint8', length: 'packagePartsLength' })
    .uint32le('permissionsLength')
    .array('permissions', { type: 'uint8', length: 'permissionsLength' })
    .uint32le('metadataLength')
    .array('metadata', { type: 'uint8', length: 'metadataLength' })
    .uint32le('permissionBindingsLength')
    .array('permissionBindings', {
        type: 'uint8',
        length: 'permissionBindingsLength',
    });

/**
 * This struct matches the metadata stream contained within the top-level binary stream.
 *
 * References:
 *
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/778afc2c-02b2-4d91-aa30-52a6067b8cb9
 */
export const ParserMetadata = new Parser()
    .endianness('little')
    .uint32le('version')
    .uint32le('metadataXmlLength')
    .array('metadataXml', { type: 'uint8', length: 'metadataXmlLength' })
    .uint32le('contentLength')
    .array('content', { type: 'uint8', length: 'contentLength' });
