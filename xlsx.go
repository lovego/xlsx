package xlsx

import (
	"io"
	"net/http"
	"os"
	"path/filepath"

	"github.com/tealeg/xlsx"
)

func WriteHttp(rw http.ResponseWriter, name string, sheets ...Sheet) error {
	file, err := Generate(sheets...)
	if err != nil {
		return err
	}

	rw.Header().Add("Content-Type",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
	)
	rw.Header().Add("Content-Disposition", `attachment;filename=`+name+`.xlsx`)

	return file.Write(rw)
}

func WriteFile(path string, sheets ...Sheet) error {
	if err := os.MkdirAll(filepath.Dir(path), 0755); err != nil {
		return err
	}
	f, err := os.OpenFile(path+".xlsx", os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0644)
	if err != nil {
		return err
	}
	defer f.Close()
	return Write(f, sheets...)
}

func Write(w io.Writer, sheets ...Sheet) error {
	file, err := Generate(sheets...)
	if err != nil {
		return err
	}
	return file.Write(w)
}

func Generate(sheets ...Sheet) (*xlsx.File, error) {
	file := xlsx.NewFile()
	for i := range sheets {
		// no generate sheet when no columns or no data
		if len(sheets[i].Columns) == 0 && sheets[i].Data == nil {
			continue
		}
		if err := sheets[i].Generate(file); err != nil {
			return nil, err
		}
	}
	return file, nil
}
