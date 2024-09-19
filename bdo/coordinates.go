package bdo

import "strings"

type coordinates []int

func (coord coordinates) valueFrom(src [][][]string) string {
	return strings.TrimSpace(src[coord[0]][coord[1]][coord[2]])
}