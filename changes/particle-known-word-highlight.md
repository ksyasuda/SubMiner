type: changed
area: overlay

- Standalone suffix tokens (MeCab pos2 接尾, e.g. さん, れる) are now excluded from JLPT/frequency/N+1 annotations by default, matching how particles and interjections are treated. Cache-backed known-word highlighting still applies; override via the pos2 exclusion config if you want them annotated.
