# Offline Dublin Core schemas

Vendored on 2026-09-28 from the official dated schema distribution:
https://www.dublincore.org/schemas/xmls/qdc/2003/04/02/

Files: dc.xsd, dcterms.xsd, dcmitype.xsd. Preserve upstream annotations.
The only changes to downloaded XSDs replace the W3C xml.xsd import URL
with ../../../ISO-IEC29500-4_2016/xml.xsd, already bundled in this skill.
Dublin Core imports remain relative to this directory. Namespace URIs
are identifiers and must not be rewritten.

Dependency closure: dc -> xml; dcterms -> xml, dc, dcmitype;
dcmitype has no schema imports. No runtime downloads are required.

| File | Upstream SHA-256 | Local SHA-256 |
| --- | --- | --- |
| `dc.xsd` | `bc1d454db1bb20dfb1867a3fe0cf007775fc4afa33a5df425c12d854f049f1f6` | `5253c7eef13f8a40ee97ac588101e44ed97d837f1d98bf1956de388b38111bc3` |
| `dcterms.xsd` | `6328c8b574c166f8ebc28d07231e6ee47121e9791689b70c3afa665f22099fed` | `02ff5ddfb370271b00d9637f982e3d0797c9e5c1e861abd234a08344996bde8c` |
| `dcmitype.xsd` | `4980845d0bedcf45e707fa64eb47492ae4f95ac4ae42f5c3c68b58c0bc9e8e49` | `4980845d0bedcf45e707fa64eb47492ae4f95ac4ae42f5c3c68b58c0bc9e8e49` |
