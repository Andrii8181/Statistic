import matplotlib.pyplot as plt
import seaborn as sns

def plot_bar(data, x, y, title="Стовпчикова діаграма"):
    plt.figure(figsize=(6,4))
    sns.barplot(data=data, x=x, y=y)
    plt.title(title)
    plt.tight_layout()
    return plt

def plot_box(data, x, y, title="Box plot"):
    plt.figure(figsize=(6,4))
    sns.boxplot(data=data, x=x, y=y)
    plt.title(title)
    plt.tight_layout()
    return plt

def plot_pie(data, labels, values, title="Кругова діаграма"):
    plt.figure(figsize=(5,5))
    plt.pie(values, labels=labels, autopct='%1.1f%%')
    plt.title(title)
    return plt
