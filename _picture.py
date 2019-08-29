import vs


class PictureParameters:
    def __init__(self):
        self.pictureName = "New Picture"
        self.createSymbol = "True"
        self.withImage = "False"
        self.imageWidth = "8\""
        self.imageHeight = "10\""
        self.imagePosition = "0.3\""
        self.imageTexture = ""
        self.withFrame = "True"
        self.frameWidth = "12\""
        self.frameHeight = "14\""
        self.frameThickness = "1\""
        self.frameDepth = "1\""
        self.frameClass = "None"
        self.frameTextureScale = "1"
        self.frameTextureRotation = "0"
        self.withMatboard = "True"
        self.matboardPosition = "0.25\""
        self.matboardClass = "None"
        self.matboardTextureScale = "1"
        self.matboardTextureRotat = "0"
        self.withGlass = "False"
        self.glassPosition = "0.75\""
        self.glassClass = "None"


class PictureRecord:
    def __init__(self):
        self.artworkTitle = ""
        self.authorName = ""
        self.artworkCreationDate = ""
        self.artworkMedia = ""
        self.type = ""
        self.roomLocation = ""
        self.artworkSource = ""
        self.registrationNumber = ""
        self.authorBirthCountry = ""
        self.authorBirthDate = ""
        self.authorDeathDate = ""
        self.designNotes = ""
        self.exhibitionMedia = ""


def build_picture(parameters: PictureParameters, record: PictureRecord or None):
    active_class = vs.ActiveClass()
    vs.NameClass("Pictures")
    if record:
        creation_record = record
    else:
        creation_record = PictureRecord()

    if parameters.createSymbol == "True":
        vs.BeginSym("{} Picture Symbol".format(parameters.pictureName))
    picture = vs.CreateCustomObjectN("Picture", (0, 0), 0, False)
    if parameters.withImage:
        texture_index = vs.Name2Index(parameters.imageTexture)
        if texture_index:
            # vs.SetTextureRefN(picture, texture_index, 100, 0)
            vs.SetTextureRefN(picture, texture_index, 0, 0)
        else:
            parameters.withImage = "False"

    vs.SetRField(picture, "Picture", "PictureName", parameters.pictureName)
    vs.SetRField(picture, "Picture", "WithImage", parameters.withImage)
    vs.SetRField(picture, "Picture", "ImageWidth", parameters.imageWidth)
    vs.SetRField(picture, "Picture", "ImageHeight", parameters.imageHeight)
    vs.SetRField(picture, "Picture", "ImagePosition", parameters.imagePosition)
    vs.SetRField(picture, "Picture", "ImageTexture", parameters.imageTexture)
    vs.SetRField(picture, "Picture", "WithFrame", parameters.withFrame)
    vs.SetRField(picture, "Picture", "FrameWidth", parameters.frameWidth)
    vs.SetRField(picture, "Picture", "FrameHeight", parameters.frameHeight)
    vs.SetRField(picture, "Picture", "FrameThickness", parameters.frameThickness)
    vs.SetRField(picture, "Picture", "FrameDepth", parameters.frameDepth)
    vs.SetRField(picture, "Picture", "FrameClass", parameters.frameClass)
    vs.SetRField(picture, "Picture", "FrameTextureScale", parameters.frameTextureScale)
    vs.SetRField(picture, "Picture", "FrameTextureRotation", parameters.frameTextureRotation)
    vs.SetRField(picture, "Picture", "WithMatboard", parameters.withMatboard)
    vs.SetRField(picture, "Picture", "MatboardPosition", parameters.matboardPosition)
    vs.SetRField(picture, "Picture", "MatboardClass", parameters.matboardClass)
    vs.SetRField(picture, "Picture", "MatboardTextureScale", parameters.matboardTextureScale)
    vs.SetRField(picture, "Picture", "MatboardTextureRotat", parameters.matboardTextureRotat)
    vs.SetRField(picture, "Picture", "WithGlass", parameters.withGlass)
    vs.SetRField(picture, "Picture", "GlassPosition", parameters.glassPosition)
    vs.SetRField(picture, "Picture", "GlassClass", parameters.glassClass)
    vs.SetName(picture, parameters.pictureName)

    vs.ResetObject(picture)

    vs.NewField("Object list data", "Artwork title", creation_record.artworkTitle, 4, 0)
    vs.NewField("Object list data", "Author name", creation_record.authorName, 4, 0)
    vs.NewField("Object list data", "Artwork creation date", creation_record.artworkCreationDate, 4, 0)
    vs.NewField("Object list data", "Artwork media", creation_record.artworkMedia, 4, 0)
    vs.NewField("Object list data", "Type", creation_record.type, 4, 0)
    vs.NewField("Object list data", "Room Location", creation_record.roomLocation, 4, 0)
    vs.NewField("Object list data", "Artwork source/lender", creation_record.artworkSource, 4, 0)
    vs.NewField("Object list data", "WDFM registration number", creation_record.registrationNumber, 4, 0)
    vs.NewField("Object list data", "Author birth country", creation_record.authorBirthCountry, 4, 0)
    vs.NewField("Object list data", "Author date of birth", creation_record.authorBirthDate, 4, 0)
    vs.NewField("Object list data", "Author date of death", creation_record.authorDeathDate, 4, 0)
    vs.NewField("Object list data", "Design notes", creation_record.designNotes, 4, 0)
    vs.NewField("Object list data", "Exhibition media", creation_record.exhibitionMedia, 4, 0)

    if parameters.createSymbol == "True":
        vs.EndSym()
        vs.SetSymbolOptionsN("{} Picture Symbol".format(parameters.pictureName), 1, 4, "Pictures")
        symbol_handle = vs.GetObject("{} Picture Symbol".format(parameters.pictureName))
        vs.Record(symbol_handle, "Object list data")
        vs.SetName(picture, parameters.pictureName)
    else:
        vs.Record(picture, "Object list data")

    vs.NameClass(active_class)
