from typing import TypedDict

class Note(TypedDict):
    number: str
    type: str
    text: str
    systemStatus: str
    userStatus: str
    equipamentNumber: str
    serieNumber: str
    material: str
    affectedAnlage: str
    damageSympton: str
    damageCode: str
    itemText: str
    isOpen: bool