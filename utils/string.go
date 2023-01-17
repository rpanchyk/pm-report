package utils

import (
	"encoding/json"
	"fmt"
	"log"
	"strings"
)

func ToPrettyString(prefix string, obj interface{}) string {
	pretty, err := json.MarshalIndent(obj, "", "  ")
	if err != nil {
		log.Fatal(err)
		return "error while prettifying"
	}
	return fmt.Sprintf("%s: \r\n%s", prefix, string(pretty))
}

func ToList(value string) []string {
	var result []string
	for _, item := range strings.Split(value, ",") {
		result = append(result, strings.Trim(item, " "))
	}
	return result
}
