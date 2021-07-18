class rel:
    def __init__(self, id: str, type: str, target: str):
        self.id = id
        self.type = type
        self.target = target

    def __str__(self):
        return f"<Relationship Id=\"{self.id}\" Type=\"{self.type}\" Target=\"{self.target}\"/>"


class rels:
    def __init__(self, rel_list: list = None):
        self.rel_list = rel_list if rel_list else []

    def __str__(self):
        return f"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{''.join([str(r) for r in self.rel_list])}</Relationships>"