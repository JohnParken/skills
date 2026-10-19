# Offline Dublin Core schemas

Preserve the official dated schema distribution vendored on 2026-09-28.
Record its origin host as `www.dublincore.org` and origin path as
`schemas/xmls/qdc/2003/04/02/`; do not fetch schemas during validation.

Keep the three source files in
`../assets/schemas/ecma/fouth-edition/dublin-core/`: `dc.xsd`, `dcterms.xsd`,
and `dcmitype.xsd`. Preserve upstream annotations and namespace URIs.

Replace only schema import locations needed for local resolution. Use
`../../../office-standard/xml.xsd` relative to each Dublin Core XSD instead
of the upstream W3C XML schema location. Keep Dublin Core imports local.
The `office-standard` folder retains the ISO/IEC 29500-4:2016 schemas under a
portable directory name; do not change the schema vocabulary.

Preserve the dependency closure: dc -> xml; dcterms -> xml, dc, dcmitype;
dcmitype has no schema imports. Do not introduce runtime downloads.

| File | Upstream SHA-256 | Local SHA-256 |
| --- | --- | --- |
| `dc.xsd` | `bc1d454db1bb20dfb1867a3fe0cf007775fc4afa33a5df425c12d854f049f1f6` | `491a0ca687ef07b6bda9b12b15c594436778ac850caa2f14e74bfbf53ce0aa4b` |
| `dcterms.xsd` | `6328c8b574c166f8ebc28d07231e6ee47121e9791689b70c3afa665f22099fed` | `a28a605c7ffe81dc28eba33899fc60c24d931fb2f26d9efd034a526958daf965` |
| `dcmitype.xsd` | `4980845d0bedcf45e707fa64eb47492ae4f95ac4ae42f5c3c68b58c0bc9e8e49` | `4980845d0bedcf45e707fa64eb47492ae4f95ac4ae42f5c3c68b58c0bc9e8e49` |
