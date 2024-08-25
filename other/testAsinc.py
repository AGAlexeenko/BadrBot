import asyncio
async def fields():
    print('Деревня1.Начали строить поле1')
    print('Деревня1.Очередь заполнена.Ожидаю')
    await asyncio.sleep(30)
    print('Деревня1.Поле достроено. Строю Поле2')

async def houses():
    print('Деревня1.Начали строить Дом1')
    await asyncio.sleep(10)
    print('Деревня1.Дом построен')

async def fields2():
    print('Деревня2.Начали строить поле1')
    print('Деревня2.Очередь заполнена.Ожидаю')
    await asyncio.sleep(20)
    print('Деревня2.Поле достроено. Строю Поле2')

async def houses2():
    print('Деревня2.Начали строить Дом1')
    await asyncio.sleep(10)
    print('Деревня2.Дом построен')
    await asyncio.sleep(5)
    print('Деревня2.Дом2 построен')
    await asyncio.sleep(5)
    print('Деревня2.Дом3 построен')



async def main():
    print('Бот запущен')
    await asyncio.gather(fields(), houses(), fields2(), houses2())
    print('Бот отработал')

asyncio.run(main())
