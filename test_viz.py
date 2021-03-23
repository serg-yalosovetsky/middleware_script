import asyncio

async def io_task():
    await asyncio.sleep(0.01)

async def main():
    t1 = asyncio.create_task(io_task())
    t2 = asyncio.create_task(io_task())
    t3 = asyncio.create_task(io_task())

    await t1
    await t2
    await t3



if __name__ == "__main__":
    asyncio.run(main())
    
    
    
    
    
    
    
    
    
    
    
    


def one():
    x = ['one', 'two']
    def inner():
        print(x)
        print(id(x))
    return inner

o = one()
o()
o.__closure__
a = o.__closure__[0].cell_contents
id(a)
dir(o.__closure__[0].cell_contents)
a.append('asdf')
a
    
    
import logging
logging.basicConfig(level='DEBUG', filename='mylog.log')   
logger = logging.getLogger()
print(logger)
print()
    
    
print(logger.level)
print()

print(logger.handlers)
name = 'serg'
def main():
    logger.debug(f'enter in the main function: name = {name}')
    
main()    





def parse_2_dict(*args):
    s = ''
    i=1
    last = ''
    if len(args)==1:
        return args[0]
    for a in args:
        if i==1:
            last =a
            i+=1
            continue
        s += '{"'+last+'":'
        last =a
        i+=1
    s += '"'+last+'"'
    for j in range(i-1):
        s+='}'
    return s

s = parse_2_dict('a', 'q','z')
s1 = parse_2_dict('s', 'w','x')
s2 = parse_2_dict('d', 'e','c','v')
repr(s)
import json
json.loads(s)
