class BaseFormatter:
    def __init__(self, doc, config):
        self.doc = doc
        self.config = config
        self.global_config = config.get('global', {})
        self.recognition_config = config.get('recognition', {})

    def format(self):
        raise NotImplementedError("子类必须实现format方法")

    def get_english_font(self):
        return self.global_config.get('english_number_font', 'Times New Roman')
