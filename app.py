import sys

if __name__ == '__main__':
    if len(sys.argv) == 3 and sys.argv[1] == '--word-worker':
        from localpdf.core import word_worker
        word_worker(sys.argv[2])
    elif len(sys.argv) == 3 and sys.argv[1] == '--verify-build':
        from localpdf.build_check import run
        sys.exit(run(sys.argv[2]))
    else:
        from localpdf.gui import main
        main()
