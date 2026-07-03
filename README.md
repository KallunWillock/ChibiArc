
![ChibiArc logo](logo_chibiarc_small.png)
# ChibiArc

![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![Platform: Windows](https://img.shields.io/badge/Platform-Windows-blue.svg)
![VBA](https://img.shields.io/badge/VBA-64bit-purple.svg)
![Dependencies](https://img.shields.io/badge/Dependencies-archiveint.dll-orange.svg)

ChibiArc is a single-class VBA module for reading and writing archive files — ZIP, 7-Zip, TAR variants, and more — directly from Excel, Word, or any Office application.

*   **Multi-Format Support:** It reads and writes a myriad of file formats (some of which I've never heard of), but importantly, this includes read/write ZIP, 7-Zip and TAR. `libarchive` will read and write ISO, but ChibiArc currently only reads ISO files. CAB and RAR are read-only under `libarchive`, but I can add code for the writing of CAB files if needed.
*   **AES Encryption:** Courtesy of the great work of WQWeto, I have adapted his encryption work and shoehorned it into the class to allow for AES-128, AES-192, and AES-256 encryption.
*   **High Performance:** Leverages `libarchive` for compression/decompression and raw Win32 APIs for I/O. The original version of this class provided both a pure VBA path and a `zlib1.dll`/`zlibwapi.dll` path, but I rewrote the whole class to rely solely on `archiveint.dll` because it is more performant.
*   **Unicode Safe:** Full Unicode support for filenames and paths, including CJK and other non‑ASCII characters - even emojis if you're feeling adventurous.
*   **ISO Helpers:** Routines to mount and dismount ISO images as virtual drives.

> [!IMPORTANT]
> As ever, any bugs, blunders, oversights, and general acts of coding inelegance are entirely my own. Any sparks of coding brilliance very likely belong to other people.

## Basic Usage
### Creating an archive file
```vba
Dim arc As New ChibiArc
If arc.NewFile("C:\output\report.zip") Then
  arc.Add "C:\reports\summary.xlsx"       ' Single file
  arc.Add "C:\reports\charts\"            ' Entire folder (recursive)
  If Not arc.SaveFile() Then
    Debug.Print "Error: " & arc.LastError
  End If
End If
```

### Extracting files/folders from an archive file
```vba
Dim arc As New ChibiArc
If arc.OpenFile("C:\output\report.zip") Then
  Dim entries As Variant
  entries = arc.Dir("*.xlsx")        ' List contents
  arc.ExtractAll "C:\extracted\"     ' Extract everything
  arc.CloseFile
End If
```

### AES-256 Encryption
```vba
Dim arc As New ChibiArc
If arc.NewFile("C:\secure\data.zip") Then
  arc.Encryption = aeAes256
  arc.PassPhrase = "MySecretPassword" 
  arc.Add "C:\sensitive\file.docx"
  arc.SaveFile
End If
```

### Mounting an ISO image
```vba
Sub TestISO_MountDismount()
  Dim arc As New ChibiArc
  Dim drive As String
  
  drive = arc.MountISO("C:\ChibiArc_Demo\demo.iso")
  If drive <> "" Then
    Debug.Print "Mounted at " & drive
    Debug.Print "First file: " & Dir(drive & "*.*")
    arc.DismountISO
  Else
    Debug.Print "Mount failed: " & arc.LastError
  End If
End Sub
```

> [!CAUTION]
> Legal note: This is covered off in the MIT License (see below/attached/to the side), but it is worth reiterating: ChibiArc is provided entirely "as is". No warranty, express or implied, is given. If you use this in production, you do so at your own risk and with my deepest sympathy.

## A 'quick' note about encryption: ZipCrypto vs AES

This is where things get slightly nuanced. `ChibiArc` supports two encryption methods for ZIP files: ZipCrypto and AES (128/192/256). They are **not** interchangeable, and choosing between them is one of those rare cases where the maxim "bigger is better" is not necessarily true.

> [!IMPORTANT]
> Windows Explorer cannot natively extract AES-encrypted ZIPs, it can only extract from ZipCrypto-encrypted ZIP files.

ZipCrypto is the original ZIP encryption. It has been around since the beginning of time(ish), and most (all?) ZIP tools can work with it. Importantly, this means Windows Explorer, which is frankly the main (and probably only) reason I've kept it in. If you need to send an encrypted archive to someone who might not have 7-Zip or WinZip installed (which in my industry, is very plausible), and you just need some level of protection, ZipCrypto will at least ensure they can open it without installing additional software.

> [!WARNING]
> ZipCrypto is apparently cryptographically weak by modern standards. I, myself, have not tested it, but I am reliably informed by the entirety of the internet that it is vulnerable to known-plaintext attacks and brute-force attempts that wouldn't take very long with today's hardware. If you have an issue with this, take it up with Microsoft... they tend to be very receptive and responsive to this sort of unsolicited feedback.

AES (specifically WinZip-compatible AES-128/192/256) is what you should be using for anything that actually matters. The important trade-off, though, is that Windows Explorer cannot natively extract AES-encrypted ZIP files. The recipient will need compatible software.

Alternatively, and unsurprisingly, the recipient can use ChibiArc.

**TL;DR:** Use ZipCrypto if compatibility with Windows Explorer is non-negotiable and the stakes are low. Use AES-256 for everything else.

> [!NOTE]
> For those playing along at home, here is one point worth noting: the Windows build of archiveint.dll handles ZipCrypto end-to-end (both reading and writing), but it cannot write AES-encrypted ZIPs. For AES, ChibiArc uses an encryption pipeline (adapted from wqweto's genius VB6/VBA AES implementation) while still leveraging libarchive for DEFLATE compression. Reading AES archives works fine through libarchive regardless - it's only the write path that needed the custom VBA layer. Invariably, someone far smarter than I will point out that the method I've adopted is not best practice and/or is brittle, and they may well be right. It does, however, **seem** to work, which is (almost always) good enough for me.

## API Overview

### Write Mode
| Method | Description |
| :--- | :--- |
| `NewFile(Path, [Type], [Pass])` | Initialise a new archive |
| `Add(Source, [Dest])` | Queue files, folders, byte arrays, or in-memory data |
| `AddNew(Name, [ParentPath])` | Queue a new empty entry for later writing |
| `SaveFile()` | Write queued items to disk and close |
| `Encryption` | Set to `aeNone`, `aeZipCrypto`, `aeAes128`, `aeAes192`, or `aeAes256` |
| `PassPhrase` | Password used for AES or ZipCrypto encryption |
| `CompressionLevel` | Set compression level (0-9) - 6 is a sensible default |
| `ConflictMode` | Choose how duplicate entry names are handled: Reject, Overwrite, or Rename |
| `SkipCompressedFileTypes` | Skip a second compression pass for already-compressed files |

### Read Mode
| Method | Description |
| :--- | :--- |
| `OpenFile(Path, [Pass])` | Open an existing archive |
| `Dir([Pattern])` | List entries - allows for wildcard patterns |
| `Extract(Name, [Path])` | Extract a single entry to disk or memory |
| `ExtractAll(Path)` | Extract all entries to a folder |
| `CloseFile()` | Close the archive handle |

### ISO helpers
| Method | Description |
| :--- | :--- |
| `MountISO(Path, [ReadOnly], [AutoDriveLetter])` | Mount an ISO image as a virtual drive |
| `MountISOToLetter(Path, DriveLetter, [ReadOnly])` | Mount an ISO image to a specific drive letter |
| `DismountISO([Force])` | Detach the currently mounted ISO |
| `ListMountedISOs()` | Returns a list of currently mounted ISO paths |
| `IsISOMounted([Path])` | Check whether an ISO image is currently mounted |

### Properties
| Property | Description |
| :--- | :--- |
| `FileCount` | Number of files in the archive |
| `FolderCount` | Number of folders in the archive |
| `CompressedSize` | Total compressed size of all entries |
| `UncompressedSize` | Total uncompressed size of all entries |
| `CompressionRatio` | Ratio of compressed vs uncompressed size |
| `IsEncrypted` | True if the archive contains encrypted entries |
| `ConflictMode` | How duplicate entry names are handled during Add: Reject, Overwrite, or Rename |
| `LastError` | Detailed error message from the last operation |
| `LastOperationTime` | Duration of the last SaveFile/ExtractAll operation (in seconds) |
| `LogLevel` | Logging verbosity |
| `ProgressInterval` | Number of entries between `ProgressEvent` notifications. The default is to raise the event for every entry |
| `AutoDismountOnTerminate` | Automatically dismount an ISO when the object is released or otherwise destroyed |

### Events
| Event | Description |
| :--- | :--- |
| `ProgressEvent` | Raised periodically during long operations (compression, extraction, etc.) |
| `IsoMounted` | Raised when an ISO image is mounted |
| `IsoDismounted` | Raised when an ISO image is dismounted |

> [!NOTE]
> This is not an exhaustive list; it's just slightly truncated because I'm exhausted writing it.

## Limitations
* **No Zip64 support** - Effectively, this caps maximum archive size at roughly 2 GB depending on entry layout. I personally have never seen a zip file even close to that size, and I query whether one should/could rely on VBA for anything of the sort.

* **ISO naming is constrained** - ChibiArc can now write ISO images, but ISO9660/Joliet rules remain restrictive. An entry with a name entirely in non-ASCII (e.g.: a Chinese/Japanese/Korean script-only folder name) is now rejected with a clear error rather than being silently dropped, so you do not lose data without knowing about it. Mixed names seem to be fine. Use ZIP or 7-Zip if you need unrestricted non-ASCII naming.

* **Windows‑only** - ChibiArc depends on `archiveint.dll`, the Windows implementation for libarchive. It ships with later versions of Windows 10+ apparently, but is not available on macOS. That said, ChibiArc is the current incarnation of a pure-VBA version, so I may eventually dust off and update that version and add that to the repo. It will rely on Cristian Buse's VBA-MemoryTools code (https://github.com/cristianbuse/VBA-MemoryTools).

* **AES requires compatible tools** - Again, Windows Explorer cannot extract AES‑encrypted ZIPs. Recipients will need 7‑Zip, WinZip, or **ChibiArc**.

* **RAR and CAB are read‑only** - Libarchive can read RAR/RAR5 and CAB, but cannot write them. CAB write support can theoretically be added if needed.

---

## Credits

ChibiArc was created by Kallun Willock (me).

* However, the AES encryption element would not be possible without **WQWeto**: <https://www.vbforums.com/showthread.php?862103>
* Indeed, I would not even know about `archiveint.dll` at all were it not for WQWeto's GZIP solution:  
<https://www.vbforums.com/showthread.php?894163-VB6-Decompress-gzip-stream-with-libarchive-on-Win10>
* I would commend his ZipArchive solution to 32-bit users: <https://github.com/wqweto/ZipArchive/>
* I was also inspired by **Cristian Buse's Excel ZipTools** repo: <https://github.com/cristianbuse/Excel-ZipTools/>
* I also made the mascot image myself, entirely in the graphical powerhouse that is Excel.
  
---

## References / resources

* LibArchive: <https://github.com/libarchive/libarchive/>
* PKWARE APPNOTE (ZIP File Format Specification): <https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT>  
  The canonical reference for ZIP structures, headers, Zip64, and traditional ZipCrypto.
* WinZip AES encryption specification: <http://www.winzip.com/aes_info.htm>

---

## Changelog

* 1.1: Added ISO write support and ISO mount/dismount helpers
* 1.0: Initial public release

## License

MIT License