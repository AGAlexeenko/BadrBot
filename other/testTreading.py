import time
from threading import Thread


def fields(x):
    print(f'Деревня{x}.Начали строить поле1')
    print(f'Деревня{x}.Очередь заполнена.Ожидаю')
    time.sleep(30)
    print(f'Деревня{x}.Поле достроено. Строю Поле2')


def houses():
    print('Деревня1.Начали строить Дом1')
    time.sleep(10)
    print('Деревня1.Дом построен')


def fields2():
    print('Деревня2.Начали строить поле1')
    print('Деревня2.Очередь заполнена.Ожидаю')
    time.sleep(20)
    print('Деревня2.Поле достроено. Строю Поле2')


def houses2():
    print('Деревня2.Начали строить Дом1')
    time.sleep(10)
    print('Деревня2.Дом построен')
    time.sleep(5)
    print('Деревня2.Дом2 построен')
    time.sleep(5)
    print('Деревня2.Дом3 построен')


def func1(i, n):
    time.sleep(3)
    print(f'В деревнe {i} строится постройка {n}')


i = 1
j = 2
N = 5


# thread1 = Thread(target=fields, args=(1,), daemon=True)
# thread2 = Thread(target=houses)
# thread3 = Thread(target=fields2)
# thread4 = Thread(target=houses2)
def main():
    threads = []
    threads1 = [Thread(target=func1, args=(i, N)) for N in range(N)]
    #threads2 = [Thread(target=func1, args=(j, N)) for N in range(N)]
    threads.extend(threads1)
    #threads.extend(threads2)


    for thread in threads:
        thread.start()

    for thread in threads:
        thread.join()

if __name__ == '__main__':
    main()
    # print('Бот запущен')
    # thread1.start()
    # thread2.start()
    # thread3.start()
    # thread4.start()
    #
    # thread2.join()
# print('Бот отработал')
