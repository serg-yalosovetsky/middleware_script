import argparse
parser = argparse.ArgumentParser(description='Great Description')

parser.add_argument('-n', action ='store', dest='n', help='simple value')
parser.add_argument('-o','--optional', type=int, default=2, help='provide an integer (default: 2)')
args = parser.parse_args()
print(args.n)
print(args.optional)