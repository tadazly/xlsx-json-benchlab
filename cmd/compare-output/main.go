package main

import (
	"crypto/md5"
	"encoding/hex"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"sync"
	"time"
)

func main() {
	left := flag.String("left", "", "left file or directory")
	right := flag.String("right", "", "right file or directory")
	maxConcurrency := flag.Int("concurrency", 0, "maximum concurrent hashing for directory input, defaults to logical CPU cores")
	flag.IntVar(maxConcurrency, "j", 0, "shorthand for -concurrency")
	flag.Parse()

	if *left == "" && flag.NArg() > 0 {
		*left = flag.Arg(0)
	}
	if *right == "" && flag.NArg() > 1 {
		*right = flag.Arg(1)
	}
	if *left == "" || *right == "" {
		log.Fatal("missing paths: use -left file|dir -right file|dir, or compare-output file|dir file|dir")
	}

	start := time.Now()
	differences, err := compare(*left, *right, *maxConcurrency)
	if err != nil {
		log.Fatal(err)
	}

	fmt.Printf("Finished in %s\n", time.Since(start).Round(time.Millisecond))
	if differences > 0 {
		os.Exit(1)
	}
}

type fileDigest struct {
	relPath string
	path    string
	size    int64
	md5     string
	err     error
}

type difference struct {
	relPath string
	left    *fileDigest
	right   *fileDigest
}

func compare(leftPath, rightPath string, maxConcurrency int) (int, error) {
	leftInfo, err := os.Stat(leftPath)
	if err != nil {
		return 0, fmt.Errorf("stat left path: %w", err)
	}
	rightInfo, err := os.Stat(rightPath)
	if err != nil {
		return 0, fmt.Errorf("stat right path: %w", err)
	}
	if leftInfo.IsDir() != rightInfo.IsDir() {
		return 0, fmt.Errorf("path types differ: left is directory=%t, right is directory=%t", leftInfo.IsDir(), rightInfo.IsDir())
	}
	if !leftInfo.IsDir() {
		return compareFiles(leftPath, rightPath)
	}
	return compareDirectories(leftPath, rightPath, maxConcurrency)
}

func compareFiles(leftPath, rightPath string) (int, error) {
	leftDigest, err := digestFile(leftPath, filepath.Base(leftPath))
	if err != nil {
		return 0, err
	}
	rightDigest, err := digestFile(rightPath, filepath.Base(rightPath))
	if err != nil {
		return 0, err
	}

	if leftDigest.size == rightDigest.size && leftDigest.md5 == rightDigest.md5 {
		fmt.Println("Files are identical.")
		fmt.Printf("left : %s size=%d md5=%s\n", leftPath, leftDigest.size, leftDigest.md5)
		fmt.Printf("right: %s size=%d md5=%s\n", rightPath, rightDigest.size, rightDigest.md5)
		return 0, nil
	}

	fmt.Println("Files differ.")
	fmt.Printf("left : %s size=%d md5=%s\n", leftPath, leftDigest.size, leftDigest.md5)
	fmt.Printf("right: %s size=%d md5=%s\n", rightPath, rightDigest.size, rightDigest.md5)
	return 1, nil
}

func compareDirectories(leftDir, rightDir string, maxConcurrency int) (int, error) {
	leftFiles, err := collectDigests(leftDir, maxConcurrency)
	if err != nil {
		return 0, fmt.Errorf("scan left directory: %w", err)
	}
	rightFiles, err := collectDigests(rightDir, maxConcurrency)
	if err != nil {
		return 0, fmt.Errorf("scan right directory: %w", err)
	}

	relativePaths := make(map[string]struct{}, len(leftFiles)+len(rightFiles))
	for path := range leftFiles {
		relativePaths[path] = struct{}{}
	}
	for path := range rightFiles {
		relativePaths[path] = struct{}{}
	}

	diffs := make([]difference, 0)
	for path := range relativePaths {
		left := leftFiles[path]
		right := rightFiles[path]
		if left == nil || right == nil || left.size != right.size || left.md5 != right.md5 {
			diffs = append(diffs, difference{relPath: path, left: left, right: right})
		}
	}
	sort.Slice(diffs, func(i, j int) bool {
		return diffs[i].relPath < diffs[j].relPath
	})

	fmt.Printf("Left files : %d (%s)\n", len(leftFiles), leftDir)
	fmt.Printf("Right files: %d (%s)\n", len(rightFiles), rightDir)
	fmt.Printf("Different files: %d\n", len(diffs))

	for _, diff := range diffs {
		fmt.Printf("\n%s\n", diff.relPath)
		printSide("left ", diff.left)
		printSide("right", diff.right)
	}

	return len(diffs), nil
}

func collectDigests(root string, maxConcurrency int) (map[string]*fileDigest, error) {
	paths, err := collectFiles(root)
	if err != nil {
		return nil, err
	}
	concurrency := optimalConcurrency(maxConcurrency, len(paths))

	jobs := make(chan string)
	results := make(chan fileDigest, len(paths))

	var wg sync.WaitGroup
	for index := 0; index < concurrency; index++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for path := range jobs {
				relPath, err := filepath.Rel(root, path)
				if err != nil {
					results <- fileDigest{path: path, err: fmt.Errorf("make relative path: %w", err)}
					continue
				}
				relPath = filepath.ToSlash(relPath)
				digest, err := digestFile(path, relPath)
				if err != nil {
					digest.err = err
				}
				results <- digest
			}
		}()
	}

	go func() {
		for _, path := range paths {
			jobs <- path
		}
		close(jobs)
		wg.Wait()
		close(results)
	}()

	digests := make(map[string]*fileDigest, len(paths))
	for result := range results {
		if result.err != nil {
			return nil, result.err
		}
		digest := result
		digests[result.relPath] = &digest
	}
	return digests, nil
}

func collectFiles(root string) ([]string, error) {
	var paths []string
	if err := filepath.WalkDir(root, func(path string, entry os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if entry.IsDir() {
			return nil
		}
		info, err := entry.Info()
		if err != nil {
			return err
		}
		if info.Mode().IsRegular() {
			paths = append(paths, path)
		}
		return nil
	}); err != nil {
		return nil, err
	}
	sort.Strings(paths)
	return paths, nil
}

func digestFile(path, relPath string) (fileDigest, error) {
	file, err := os.Open(path)
	if err != nil {
		return fileDigest{}, fmt.Errorf("open %s: %w", path, err)
	}
	defer func() {
		if closeErr := file.Close(); closeErr != nil {
			log.Printf("close %s: %v", path, closeErr)
		}
	}()

	stat, err := file.Stat()
	if err != nil {
		return fileDigest{}, fmt.Errorf("stat %s: %w", path, err)
	}

	hash := md5.New()
	if _, err := io.Copy(hash, file); err != nil {
		return fileDigest{}, fmt.Errorf("hash %s: %w", path, err)
	}

	return fileDigest{
		relPath: relPath,
		path:    path,
		size:    stat.Size(),
		md5:     hex.EncodeToString(hash.Sum(nil)),
	}, nil
}

func printSide(label string, digest *fileDigest) {
	if digest == nil {
		fmt.Printf("  %s: missing\n", label)
		return
	}
	fmt.Printf("  %s: size=%d md5=%s\n", label, digest.size, digest.md5)
}

func optimalConcurrency(maxConcurrency int, taskCount int) int {
	concurrency := runtime.NumCPU()
	if maxConcurrency > 0 && maxConcurrency < concurrency {
		concurrency = maxConcurrency
	}
	if concurrency < 1 {
		concurrency = 1
	}
	if taskCount > 0 && concurrency > taskCount {
		concurrency = taskCount
	}
	return concurrency
}
