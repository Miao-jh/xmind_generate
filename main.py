import xmind
from xmind.core.topic import TopicElement
import pandas as pd
import os


class MakeXmind:

    def __init__(self, save_to, data_from, title):
        self.save_to = save_to
        self.data_from = data_from
        self.title = title
        self.data = pd.DataFrame()
        self.deep = 0

    def read_data(self):
        self.data = pd.read_excel(self.data_from)
        self.deep = self.data.shape[1]

    def make(self):
        self.read_data()
        if os.path.exists(self.save_to):
            os.remove(self.save_to)
        workbook = xmind.load(self.save_to)
        sheet = workbook.getPrimarySheet()
        sheet.setTitle(self.title)
        root_topic = sheet.getRootTopic()
        root_topic.setTitle(self.title)
        self.add_sub(root_topic, 0)
        xmind.save(workbook, path=self.save_to)

    def add_sub(self, root_topic: TopicElement, level, root_name=None):
        if root_name is None and level == 0:
            deep_i_topic = self.data.iloc[:, level].value_counts()
        else:
            deep_i_topic = self.data.loc[self.data.iloc[:, level-1] == root_name,
                                         self.data.columns[level]].value_counts()
        for sub_name, sub_count in deep_i_topic.iteritems():
            sub_title = str(sub_name) + "_" + str(sub_count)
            sub_topic = root_topic.addSubTopic()
            sub_topic.setTitle(sub_title)
            if level < self.deep-1:
                self.add_sub(sub_topic, level+1, root_name=sub_name)


if __name__ == '__main__':
    x = MakeXmind('test.xmind', data_from='test.xlsx', title='用户信息')
    x.make()
