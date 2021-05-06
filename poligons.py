import matplotlib.pyplot as plt

poligon = [[0, 0], [0, 10], [5, 10], [5, 15], [10, 15], [10, 8], [
    17, 8], [17, 17], [20, 17], [20, 14], [25, 14], [25, 0], [0, 0]]
parallel = [[0, 0], [0, 25]]


def plot_poligon(poligon_coordinates):
    xs, ys = zip(*poligon)
    plt.figure()
    plt.plot(xs, ys)
    plt.xlabel("x")
    plt.ylabel("y")


def plot_parallel(parallel):
    xs, ys = zip(*parallel)
    plt.plot(xs, ys, "r--")


plot_poligon(poligon)
plot_parallel(parallel)


plt.show()
