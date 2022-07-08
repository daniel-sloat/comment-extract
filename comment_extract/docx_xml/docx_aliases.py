from typing import TypeAlias
from datetime import datetime


# Example StyleProps: {"b": {}, "u": {"val: "true"}}
# Example Styles: {"Heading1": {"b": {}, "u": {"val: "true"}}}
StyleProps: TypeAlias = dict[str : dict[str:str]]
Styles: TypeAlias = dict[str:StyleProps]


Paragraphs: TypeAlias = list
Runs: TypeAlias = list
RunText: TypeAlias = str
RunProps: TypeAlias = StyleProps  # dict[str : dict[str:str]]
EncodedRunProps: TypeAlias = str
# //TODO Do run props need to be encoded?
Run: TypeAlias = list[RunText, RunProps | EncodedRunProps]
NestedRunData: TypeAlias = Paragraphs[Runs[Run]]
DataLinkedToComment: TypeAlias = dict[int:NestedRunData]
CommentRecordData: TypeAlias = dict[str : str | datetime | NestedRunData | None]
