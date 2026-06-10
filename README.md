# Solidworks-Open

Opens a SolidWorks file by job number: the drawing if one exists, otherwise the assembly.

## Usage

Run `SolidworksOpen.bas` from SolidWorks. A dialog asks for the job number, then the macro
locates the job folder and opens the first file it finds, in this order:

1. `<job>-01.SLDDRW`
2. `<job>-02.SLDDRW`
3. `<job>-01.SLDASM`
4. `<job>-02.SLDASM`

If the file is already open in the current SolidWorks session, it is brought to the front
instead of reopened.

## Folder hierarchy

The job folder is found under `Z:\Solidworks\Current\JOBS\` using the same hierarchy as
[Pack-n-Go](../Pack-n-Go), probing each job type in turn:

| Job type     | Intermediate folder                              | Example                            |
| ------------ | ------------------------------------------------ | ---------------------------------- |
| GENERAL LINE | Range of first 3 digits in groups of 5           | `GENERAL LINE\406-410\406123\`     |
| HD-PFD       | First 2 digits + `XXXX`                          | `HD-PFD\40XXXX\406123\`            |
| HDX          | Range of first 3 digits in groups of 5           | `HDX\406-410\406123\`              |
| AXIAL        | None (job sits directly under `AXIAL\`)          | `AXIAL\406123\`                    |

Range folders special-case `401-405` as `400-405`, matching Pack-n-Go.
