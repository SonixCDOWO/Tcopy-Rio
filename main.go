package main

import (
	"bytes"
	"fmt"
	"sort"
	"strconv"

	"html/template"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
	"sync" // Agregado para el manejo seguro de archivos en entornos concurrentes

	"encoding/json"
	"net/http"
	"os/exec"
	"regexp"
	"runtime"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	wkhtml "github.com/SebastiaanKlippert/go-wkhtmltopdf" // Para PDF
)

// --- CONFIGURACIÓN DROPBOX ---
// DROPBOX TOKEN ya no se usa
const DROPBOX_TOKEN = "sl.u.AGQXLy0X9LTgzxU6MVV2JB3wrJvSF8FXOrzGxz4sw5u7uO9TGizc7QY4M705qtl_AhebScr_74iB8tulXuQF93-uPEF3T0YRWKcLnhtWJcDEjNVePyfSPtsPwiwragvPYa3j61AbASO5a1jI8gGt4iIkCcj7UDDPxxju4rEcga-WUIET1TeTc1kFBNplncc_nM9kd0K7df5Ai6bjyUhMHaNT7WTUoYVMCkQ1mCYbHu8HzDufTFoUr2yNo7ldyqyUb6HnysNaeF4LbEZ_TnK3wsuQOQ10x-18Gpn0-Ynmz08IPoVujlKus_IpKOCZiBVzpfpCzhd8iMA83sOVqMuTa0YLDqi9syOz7X5fq_wdIvhAFxzG3nnskioco6GPRlQEoYvzzUEJjtvwWa8iDX5jyUgIA89dMSTBTXKUcp9ia4h8NaDVWGFtMaVR2YJpLzsKRNy0bqiFcUI8ej_ITuaQPNJeQ5KRx67kSDrWXNfTYJHMl5_wX8Lofs0C0Twm_8RVuVlsycDy31eUEqJILapEE8QuZuuHMp5utustEtkrwYZR7T-9m--OcPv2C4Laq4ZBZn0cPxvzQJqxBXHUtAATugGktcok3E5fiq_swdnVAY-cnzFCVIV-OAaFVKmCxkzSR1VdkiBfYGsswEks4xK9TxXz3QEZYKUzNKEMtJNwsFqWoAhwg4SNf5fCAEBeYL8Vhwxd-aYE_FEuWpWi3BrBQNeFR5qV59nuhQTkSNRQCZAq_YSc7GMWCbMUBdIDAF0-W_AtpJfpbVj42GLQGvuDfwZX-5JUSM-KAZKQNBQScCHxcyWFT8ZQYlh_G_UbuzAbO40Ced6bO3zitdcBf_r8CsHQEfzKbj-d3kT1iAb3V9nwF536VXKRKd2CqfiAx6d4egK365goqDaomFoXmgnApkfHzufQ2jdpFmu2Bvvy2EqRzqZLaeOfuGyVnGDaa12sK-99o-dqD2VR0seRC0PYJYYwaZRs3ZSuL9YLrc7ej5fb-AAf38L7IbH8rr3J-lRqWL_jwXW8MQr0UMMnVp2tBXf6rhQpV362TGtziTb-chSX9ZW_Rbjg_Y5d9wcTxrdDjjIlvyo_jVnO5BK9bQ9qEySohxB9hlweJL8XoDJLOPrDGjJBoF83EeyEOymeMZbrrb2JuhuRHGj9tjOusZW7j8WVKFwlDPiXRd0S7Yni4lRZn9nKS20yHfpxZZ2bPyYLrwN6YZt1rRSvDQTQwHSw2v9eMUVXnGAKDRlT2elm7odUzIhY7G5nXq7BrA93suJDHzohKgdRfw5N-3I_1CEELg_kLyrRA_tsGZO4xhRx9EqIZpysrwEPMSvOqhdqwcKIo1oSq8Y9vEXSNQCPMvHZFpuh6hWi1B853Wd1mqVdQkl2IiF8Kv2AC-CBM3cjP9hgPWpred3-Rgxx9dqpQznMAa2qooBQJhUze0_aPJfc_cHpyTF2DYF0ifIZoVqFes2UiCw" // El token que generaste en Dropbox Developers
const DROPBOX_FILE_PATH = "/CENSO GENERAL NUEVO.xlsx"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            // Ruta dentro de Dropbox
// -----------------------------

// Variables globales nuevas
const APP_KEY = "a2yn025pt265yjg"
const APP_SECRET = "z0k5469on50e9lu"
const REFRESH_TOKEN = "zNV6Kp0-xlUAAAAAAAAAAQIk0QMq9TGPk3ApbkRlmp8-TmUg_U-slcOmVnCYMg1C"

// Función para obtener un token de acceso válido usando el Refresh Token
func obtenerAccessToken() (string, error) {
	url := "https://api.dropbox.com/oauth2/token"
	data := "grant_type=refresh_token&refresh_token=" + REFRESH_TOKEN

	req, _ := http.NewRequest("POST", url, strings.NewReader(data))
	req.SetBasicAuth(APP_KEY, APP_SECRET)
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")

	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()

	var result map[string]interface{}
	json.NewDecoder(resp.Body).Decode(&result)

	return result["access_token"].(string), nil
}

// Estructura para mapear la respuesta de Dropbox al descargar
type DropboxMetadata struct {
	PathDisplay string `json:"path_display"`
	Rev         string `json:"rev"` // Aquí viene el "Commit Hash"
}

// Excel
const EXCEL_FILE = "CENSO GENERAL NUEVO.xlsx"
const PRIMERA_HOJA = "CENSO"

const HISTORY_FILE = "history.json"
const ACTIVITIES_FILE = "activities.json"

var (
	ultimoRevGuardado string
	excelMutex        sync.Mutex // Mutex para coordinar la lectura y escritura segura del archivo físico Excel
)

// Carga los logs desde el archivo al iniciar el programa
func loadLogsFromFile() {
	if _, err := os.Stat(HISTORY_FILE); os.IsNotExist(err) {
		return // Si el archivo no existe, no hace nada
	}
	data, err := ioutil.ReadFile(HISTORY_FILE)
	if err != nil {
		fmt.Println("Error al leer el archivo de historial:", err)
		return
	}
	json.Unmarshal(data, &historyLogs)
}

type HistoryEntry struct {
	User        string    `json:"user"`
	Description string    `json:"description"`
	Timestamp   time.Time `json:"timestamp"`
}

// Guarda los logs en el archivo cada vez que hay un cambio
func saveLogsToFile() {
	data, err := json.MarshalIndent(historyLogs, "", "  ")
	if err != nil {
		fmt.Println("Error al codificar el historial:", err)
		return
	}
	err = ioutil.WriteFile(HISTORY_FILE, data, 0644)
	if err != nil {
		fmt.Println("Error al escribir el archivo de historial:", err)
	}
}

var historyLogs []HistoryEntry

func addLog(description string) {
	entry := HistoryEntry{
		User:        "Operador Dropbox",
		Description: description,
		Timestamp:   time.Now(),
	}
	// Agregamos al inicio del slice
	historyLogs = append([]HistoryEntry{entry}, historyLogs...)

	// GUARDAR EN DISCO
	saveLogsToFile()
}

// Define qué columnas usarán una búsqueda "Contains" (contiene).
// Las columnas que NO estén en este mapa usarán una búsqueda exacta (==).
var containsSearchColumns = map[string]bool{
	"Nombre completo":     true,
	"Cedula de identidad": true,
}

type Filter struct {
	Column string `json:"column"`
	Value  string `json:"value"`
}

// Galeria
type ImageInfo struct {
	Filename   string `json:"filename"`
	UploadDate string `json:"upload_date"`
}

type GalleryData struct {
	Images   []ImageInfo
	Messages []string
}

// Estructura para un nodo en la vista de árbol (jerarquía)
type TreeNode struct {
	Text     string      `json:"text"`
	Type     string      `json:"type"`
	Children []*TreeNode `json:"children"`
	State    struct {
		Opened bool `json:"opened"`
	} `json:"state"`
}

// Estructura para una persona individual
type Person struct {
	Parentesco string `json:"parentesco"`
	Nombres    string `json:"nombres"`
	Documento  string `json:"documento"`
}

const uploadDir = "assets/imagenes"

// Abre la URL en el navegador predeterminado del sistema operativo.
func openBrowser(url string) {
	var err error

	switch runtime.GOOS {
	case "windows":
		err = exec.Command("rundll32", "url.dll,FileProtocolHandler", url).Start()
	case "darwin":
		err = exec.Command("open", url).Start()
	case "linux":
		err = exec.Command("xdg-open", url).Start()
	default:
		fmt.Println("No se pudo abrir el navegador automáticamente.")
	}

	if err != nil {
		fmt.Println("Error al abrir el navegador:", err)
	}
}

// Estructura de respuesta de metadatos de archivos de la API de Dropbox
type DropboxMetadataResponse struct {
	Name           string `json:"name"`
	Rev            string `json:"rev"`
	Size           int64  `json:"size"`
	ClientModified string `json:"client_modified"`
}

// obtenerMetadataRemota consulta la metadata oficial del archivo en Dropbox
func obtenerMetadataRemota() (*DropboxMetadataResponse, error) {
	token, err := obtenerAccessToken()
	if err != nil {
		return nil, err
	}

	url := "https://api.dropbox.com/2/files/get_metadata"
	bodyArgs := map[string]interface{}{
		"path":                                DROPBOX_FILE_PATH,
		"include_media_info":                  false,
		"include_deleted":                     false,
		"include_has_explicit_shared_members": false,
	}
	jsonArgs, _ := json.Marshal(bodyArgs)

	req, err := http.NewRequest("POST", url, bytes.NewBuffer(jsonArgs))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Content-Type", "application/json")

	client := &http.Client{Timeout: 8 * time.Second}
	resp, err := client.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		bodyBytes, _ := ioutil.ReadAll(resp.Body)
		return nil, fmt.Errorf("error api dropbox (status %s): %s", resp.Status, string(bodyBytes))
	}

	var meta DropboxMetadataResponse
	err = json.NewDecoder(resp.Body).Decode(&meta)
	if err != nil {
		return nil, err
	}
	return &meta, nil
}

// syncStatusHandler devuelve al frontend el estado actual de la base de datos comparada con la de Dropbox
func syncStatusHandler(w http.ResponseWriter, r *http.Request) {
	localExists := false
	var localSize int64
	var localMtime string

	if info, err := os.Stat(EXCEL_FILE); err == nil {
		localExists = true
		localSize = info.Size()
		localMtime = info.ModTime().Format("02/01/2006 03:04 PM")
	}

	remoteMeta, err := obtenerMetadataRemota()
	status := "Desconocido"
	remoteRev := ""
	remoteSize := int64(0)
	synced := false
	errConnecting := false

	if err == nil && remoteMeta != nil {
		remoteRev = remoteMeta.Rev
		remoteSize = remoteMeta.Size

		if localExists && localSize > 0 && ultimoRevGuardado == remoteMeta.Rev {
			status = "Sincronizado"
			synced = true
		} else if !localExists || localSize == 0 {
			status = "Base de datos local vacía"
		} else {
			status = "Actualización disponible en la nube"
		}
	} else {
		errConnecting = true
		if localExists && localSize > 0 {
			status = "Modo Fuera de Línea (Archivo local disponible)"
		} else {
			status = "Sin conexión (Base de datos ausente)"
		}
	}

	response := map[string]interface{}{
		"local_exists":     localExists,
		"local_size":       localSize,
		"local_mtime":      localMtime,
		"local_rev":        ultimoRevGuardado,
		"remote_rev":       remoteRev,
		"remote_size":      remoteSize,
		"status":           status,
		"synced":           synced,
		"error_connecting": errConnecting,
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(response)
}

// forceSyncHandler permite al usuario forzar la descarga del censo directamente desde la nube
func forceSyncHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Método no permitido", http.StatusMethodNotAllowed)
		return
	}

	err := descargarDeDropbox()
	if err != nil {
		http.Error(w, "Error en la descarga de sincronización: "+err.Error(), http.StatusInternalServerError)
		return
	}

	addLog("Sincronización: Se forzó la descarga de la base de datos desde la nube")

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]interface{}{
		"success": true,
		"rev":     ultimoRevGuardado,
	})
}

// ------------------- MANEJO DEL EXCEL -------------------------
// deleteRowHandler elimina una fila específica del archivo Excel.
func deleteRowHandler(w http.ResponseWriter, r *http.Request) {
	// Sincronización preventiva previa a la eliminación
	descargarDeDropbox()

	excelMutex.Lock()
	defer excelMutex.Unlock()

	fmt.Println("--- LOG: Endpoint /api/delete-row invocado de manera directa. ---")

	var req struct {
		Row int `json:"__row"`
	}

	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		fmt.Printf("--- ERROR: No se pudo decodificar el payload JSON: %v ---\n", err)
		http.Error(w, "Payload inválido", http.StatusBadRequest)
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	if err := f.RemoveRow(PRIMERA_HOJA, req.Row); err != nil {
		http.Error(w, "Error al remover la fila", http.StatusInternalServerError)
		return
	}

	if err := f.Save(); err != nil {
		http.Error(w, "No se guardó el Excel", http.StatusInternalServerError)
		return
	}

	w.WriteHeader(http.StatusOK)
	addLog("Base de Datos: Se eliminó una fila directamente")

	go subirADropbox()
}

func normalizeHeader(header string) string {
	lower := strings.ToLower(header)
	reg := regexp.MustCompile("[^a-z0-9]+")
	return reg.ReplaceAllString(lower, "")
}

func bulkImportHandler(w http.ResponseWriter, r *http.Request) {
	// 1. Sincronización preventiva
	descargarDeDropbox()

	excelMutex.Lock()
	defer excelMutex.Unlock()

	var req struct {
		Datos []map[string]string `json:"datos"`
	}
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, "Payload inválido", http.StatusBadRequest)
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	rows, _ := f.GetRows(PRIMERA_HOJA)
	if len(rows) == 0 {
		http.Error(w, "Archivo vacío", http.StatusInternalServerError)
		return
	}
	headers := rows[0]

	// --- LOGICA DE ACTUALIZACIÓN VS INSERCIÓN ---

	// 1. Mapear cabeceras para saber qué columna es cada cual
	headerMap := make(map[string]int)
	cedulaCol := -1
	for i, h := range headers {
		headerMap[h] = i
		if normalizeHeader(h) == normalizeHeader("Cedula de identidad") {
			cedulaCol = i
		}
	}

	// 2. Mapear la ubicación de cada Cédula en el archivo (Cedula -> Numero de Fila)
	// Esto nos dirá en qué fila está cada persona actualmente
	rowMap := make(map[string]int)
	for i, row := range rows {
		if i == 0 || len(row) <= cedulaCol {
			continue
		}
		ced := cleanCedula(row[cedulaCol])
		if ced != "" && !strings.Contains(strings.ToLower(ced), "menor") {
			rowMap[ced] = i + 1 // Guardamos fila (base 1)
		}
	}

	nextAvailableRow := len(rows) + 1

	for _, personaImportada := range req.Datos {
		cedImportada := cleanCedula(personaImportada["Cedula de identidad"])

		targetRow := -1

		// ¿Ya existe esta persona?
		if rowIdx, existe := rowMap[cedImportada]; existe && cedImportada != "" {
			targetRow = rowIdx // Vamos a sobreescribir su fila actual
		} else {
			targetRow = nextAvailableRow // Es nuevo, va al final
			nextAvailableRow++
		}

		// Escribir los datos en la fila seleccionada
		for key, val := range personaImportada {
			if colIdx, ok := headerMap[key]; ok {
				cell, _ := excelize.CoordinatesToCellName(colIdx+1, targetRow)

				// TRUCO: Si el valor parece un número científico (como el teléfono en tu imagen),
				// lo forzamos a texto para que no se dañe
				if strings.Contains(strings.ToLower(val), "e+") || strings.Contains(val, "E+") {
					// Limpiar formato científico si es necesario o dejar como string plano
					f.SetCellStr(PRIMERA_HOJA, cell, val)
				} else {
					f.SetCellValue(PRIMERA_HOJA, cell, val)
				}
			}
		}
	}

	if err := f.Save(); err != nil {
		http.Error(w, "No se pudo guardar el Excel", http.StatusInternalServerError)
		return
	}

	// Subir cambios a la nube
	go subirADropbox()

	addLog("Importación: Se procesó una carga masiva (Actualizaciones e Inserciones)")
	w.WriteHeader(http.StatusOK)
}

// Estructura para el resultado de la comparación
type ImportPreviewRow struct {
	Status string            `json:"status"`  // "nuevo", "igual", "cambio"
	Data   map[string]string `json:"data"`    // Lo que viene del Excel
	DbData map[string]string `json:"db_data"` // Toda la info que ya existe en DB
}

func previewImportHandler(w http.ResponseWriter, r *http.Request) {
	// 1. Recepción segura del archivo
	if err := r.ParseMultipartForm(10 << 20); err != nil {
		http.Error(w, "Archivo demasiado grande", http.StatusBadRequest)
		return
	}

	file, _, err := r.FormFile("file")
	if err != nil {
		http.Error(w, "No se encontró el archivo en la petición", http.StatusBadRequest)
		return
	}
	defer file.Close()

	fUploaded, err := excelize.OpenReader(file)
	if err != nil {
		http.Error(w, "Error al abrir el lector de Excel", http.StatusInternalServerError)
		return
	}

	// Leer filas del Excel subido (asumimos la primera hoja)
	rowsUploaded, _ := fUploaded.GetRows(fUploaded.GetSheetList()[0])
	if len(rowsUploaded) < 2 {
		json.NewEncoder(w).Encode(map[string]interface{}{"data": []ImportPreviewRow{}, "duplicados_omitidos": 0})
		return
	}
	headersUploaded := rowsUploaded[0]

	// 2. Cargar Base de Datos Local para comparar (Modo lectura)
	excelMutex.Lock()
	fLocal, _ := excelize.OpenFile(EXCEL_FILE)
	rowsLocal, _ := fLocal.GetRows(PRIMERA_HOJA)
	excelMutex.Unlock()

	headersLocal := rowsLocal[0]
	cedulaColLocal := -1
	for i, h := range headersLocal {
		if normalizeHeader(h) == normalizeHeader("Cedula de identidad") {
			cedulaColLocal = i
			break
		}
	}

	// Creamos un mapa de la DB: Cedula -> Mapa de todos sus campos
	dbMap := make(map[string]map[string]string)
	for i, row := range rowsLocal {
		if i == 0 || len(row) <= cedulaColLocal {
			continue
		}
		ced := cleanCedula(row[cedulaColLocal])
		if ced != "" {
			pData := make(map[string]string)
			for j, val := range row {
				if j < len(headersLocal) {
					pData[headersLocal[j]] = val
				}
			}
			dbMap[ced] = pData
		}
	}

	// 3. Procesar Deduplicación y Comparación detallada
	seenInExcel := make(map[string]bool)
	internalDuplicatesCount := 0
	var result []ImportPreviewRow

	// Identificar columnas clave en el Excel subido
	cedIdx, nomIdx, comIdx, torIdx, casIdx := -1, -1, -1, -1, -1
	for i, h := range headersUploaded {
		cleanH := normalizeHeader(h)
		if cleanH == normalizeHeader("Cedula de identidad") {
			cedIdx = i
		}
		if cleanH == normalizeHeader("Nombre completo") {
			nomIdx = i
		}
		if cleanH == normalizeHeader("Comunidad") {
			comIdx = i
		}
		if cleanH == normalizeHeader("Torre") {
			torIdx = i
		}
		if cleanH == normalizeHeader("Numero de casa / apto") || cleanH == "casa" {
			casIdx = i
		}
	}

	getV := func(row []string, idx int) string {
		if idx != -1 && idx < len(row) {
			return strings.TrimSpace(row[idx])
		}
		return ""
	}

	for i, row := range rowsUploaded {
		if i == 0 {
			continue
		} // Ignorar cabecera

		ced := cleanCedula(getV(row, cedIdx))
		nom := strings.TrimSpace(getV(row, nomIdx))

		// --- NUEVO FILTRO ESTRICTO ---
		// Si la cédula está vacía, se ignora la fila de inmediato. No importa si tiene nombre.
		if ced == "" {
			continue
		}

		// Además, si el nombre es demasiado corto (basura como "."), también lo ignoramos
		if len(nom) < 2 {
			continue
		}

		// FIRMA ÚNICA para evitar duplicados en el mismo archivo
		var signature string
		if ced != "" && !strings.Contains(strings.ToLower(ced), "menor") {
			signature = "ced_" + ced
		} else {
			signature = fmt.Sprintf("nom_%s_loc_%s_%s_%s",
				normalizeHeader(getV(row, nomIdx)),
				normalizeHeader(getV(row, comIdx)),
				normalizeHeader(getV(row, torIdx)),
				normalizeHeader(getV(row, casIdx)))
		}

		if seenInExcel[signature] {
			internalDuplicatesCount++
			continue
		}
		seenInExcel[signature] = true

		// Mapear los datos que vienen en la fila del Excel
		rowData := make(map[string]string)
		for j, val := range row {
			if j < len(headersUploaded) {
				rowData[headersUploaded[j]] = val
			}
		}

		// COMPARACIÓN CON LA BASE DE DATOS
		status := "nuevo"
		var dbPersonData map[string]string = nil

		if dbPerson, exists := dbMap[ced]; exists && ced != "" {
			dbPersonData = dbPerson // Enviamos la fila completa de la DB al front
			status = "igual"

			// Solo marcamos como "cambio" si hay diferencias en datos NO vacíos
			for keyExcel, valExcel := range rowData {
				valDB, ok := dbPerson[keyExcel]
				if !ok {
					continue
				} // Si la columna no existe en DB, ignorar

				vEx := strings.TrimSpace(valExcel)
				vDB := strings.TrimSpace(valDB)

				// Solo si el Excel trae algo y es diferente a lo que hay (sin importar mayúsculas)
				if vEx != "" && !strings.EqualFold(vEx, vDB) {
					status = "cambio"
					break
				}
			}
		}

		result = append(result, ImportPreviewRow{
			Status: status,
			Data:   rowData,
			DbData: dbPersonData, // Esto permite al modal mostrar campos que no están en el Excel
		})
	}

	// 4. Respuesta al cliente
	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]interface{}{
		"data":                result,
		"duplicados_omitidos": internalDuplicatesCount,
	})
}

// checkCedulasHandler recibe una lista de cédulas y devuelve las que ya existen.
func checkCedulasHandler(w http.ResponseWriter, r *http.Request) {
	// Forzar la descarga preventiva desde Dropbox
	descargarDeDropbox()

	excelMutex.Lock()
	defer excelMutex.Unlock()

	fmt.Println("--- LOG (check-cedulas): Endpoint invocado. ---")
	var req struct {
		Cedulas []string `json:"cedulas"`
	}
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, "Payload inválido", http.StatusBadRequest)
		return
	}
	fmt.Printf("--- LOG (check-cedulas): Recibidas %d Cédulas para verificar.\n", len(req.Cedulas))

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "No se pudo abrir el Excel", http.StatusInternalServerError)
		return
	}
	rows, _ := f.GetRows(PRIMERA_HOJA)

	cedulaHeaderNormalized := normalizeHeader("Cedula de identidad")
	cedulaColIndex := -1
	for i, h := range rows[0] {
		if normalizeHeader(h) == cedulaHeaderNormalized {
			cedulaColIndex = i
			break
		}
	}
	if cedulaColIndex == -1 {
		fmt.Println("--- ERROR (check-cedulas): No se encontró la columna 'Cedula de identidad' en el archivo principal.")
		http.Error(w, "No se encontró la columna de Cedula", http.StatusInternalServerError)
		return
	}

	existingCedulas := make(map[string]bool)
	for i, row := range rows {
		if i == 0 {
			continue
		}
		if cedulaColIndex < len(row) {
			cleanDBVal := cleanCedula(row[cedulaColIndex])
			existingCedulas[cleanDBVal] = true
		}
	}

	var duplicates []string
	for _, cedula := range req.Cedulas {
		cleanReqCed := cleanCedula(cedula)
		if _, exists := existingCedulas[cleanReqCed]; exists {
			duplicates = append(duplicates, cedula)
		}
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(duplicates)
}

// getPersonByCedulaHandler busca una persona por su cédula.
func getPersonByCedulaHandler(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	cedulaToFind := cleanCedula(r.URL.Query().Get("cedula"))
	fmt.Printf("--- LOG (get-person): Endpoint invocado para buscar la cédula: %s ---\n", cedulaToFind)

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "No se pudo abrir el Excel", http.StatusInternalServerError)
		return
	}
	rows, _ := f.GetRows(PRIMERA_HOJA)
	headers := rows[0]

	cedulaHeaderNormalized := normalizeHeader("Cedula de identidad")
	cedulaColIndex := -1
	for i, h := range headers {
		if normalizeHeader(h) == cedulaHeaderNormalized {
			cedulaColIndex = i
			break
		}
	}
	if cedulaColIndex == -1 {
		http.Error(w, "No se encontró la columna de Cedula", http.StatusInternalServerError)
		return
	}

	var personData map[string]string
	for _, row := range rows[1:] {
		if cedulaColIndex < len(row) {
			dbCedula := cleanCedula(row[cedulaColIndex])
			if dbCedula == cedulaToFind {
				personData = make(map[string]string)
				for j, header := range headers {
					if j < len(row) {
						personData[header] = row[j]
					}
				}
				break
			}
		}
	}

	if personData == nil {
		http.NotFound(w, r)
		return
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(personData)
}

// Estructuras para las respuestas JSON
type ExcelResponse struct {
	Headers []string            `json:"headers"`
	Data    []map[string]string `json:"data"`
}

// Estructura para la respuesta del DataTables
type DTResponse struct {
	Draw            int                 `json:"draw"`
	RecordsTotal    int                 `json:"recordsTotal"`
	RecordsFiltered int                 `json:"recordsFiltered"`
	Data            []map[string]string `json:"data"`
}

// Obtiene las columnas del Excel y las devuelve como JSON
func getColumns(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "no se pudo abrir el Excel", 500)
		return
	}

	row, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(row) == 0 {
		http.Error(w, "sheet vacío o no existe", 500)
		return
	}

	headers := row[0]
	cleanHeaders := make([]string, len(headers))
	for i, h := range headers {
		cleanHeaders[i] = strings.TrimSpace(h)
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(cleanHeaders)
}

// Leer datos del Excel y paginarlos para DataTables
// Leer datos del Excel y paginarlos para DataTables (con soporte para rango de edad)
func getData(w http.ResponseWriter, r *http.Request) {
	descargarDeDropbox()

	excelMutex.Lock()
	defer excelMutex.Unlock()

	search := strings.ToLower(r.URL.Query().Get("search[value]"))
	start, _ := strconv.Atoi(r.URL.Query().Get("start"))
	length, _ := strconv.Atoi(r.URL.Query().Get("length"))
	draw, _ := strconv.Atoi(r.URL.Query().Get("draw"))

	// Parámetros de rango de edad
	minAgeStr := r.URL.Query().Get("minAge")
	maxAgeStr := r.URL.Query().Get("maxAge")
	var minAgeVal, maxAgeVal int
	var hasMinAge, hasMaxAge bool

	if minAgeStr != "" {
		if val, err := strconv.Atoi(minAgeStr); err == nil {
			minAgeVal = val
			hasMinAge = true
		}
	}
	if maxAgeStr != "" {
		if val, err := strconv.Atoi(maxAgeStr); err == nil {
			maxAgeVal = val
			hasMaxAge = true
		}
	}

	filtersJSON := r.URL.Query().Get("filters")
	var activeFilters []Filter
	if filtersJSON != "" {
		if err := json.Unmarshal([]byte(filtersJSON), &activeFilters); err != nil {
			http.Error(w, "Filtros inválidos", http.StatusBadRequest)
			return
		}
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "no se pudo abrir el Excel", 500)
		return
	}

	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) == 0 {
		http.Error(w, "error leyendo filas", 500)
		return
	}

	keys := rows[0]

	// Buscar índice de la columna "Edad"
	edadColIndex := -1
	for j, key := range keys {
		if strings.TrimSpace(strings.ToLower(key)) == "edad" {
			edadColIndex = j
			break
		}
	}

	type IndexedRow struct {
		Index int
		Cells []string
	}

	filtered := make([]IndexedRow, 0)
	for i := 1; i < len(rows); i++ {
		row := rows[i]

		globalMatch := false
		if search == "" {
			globalMatch = true
		} else {
			for _, cell := range row {
				if strings.Contains(strings.ToLower(cell), search) {
					globalMatch = true
					break
				}
			}
		}

		multiColumnMatch := true
		for _, filter := range activeFilters {
			// Ignorar texto de filtro si viene vacío
			if strings.TrimSpace(filter.Value) == "" {
				continue
			}

			filterColumnIndex := -1
			for j, key := range keys {
				if strings.TrimSpace(key) == filter.Column {
					filterColumnIndex = j
					break
				}
			}

			if filterColumnIndex != -1 && filterColumnIndex < len(row) {
				cellValue := strings.ToLower(strings.TrimSpace(row[filterColumnIndex]))
				filterVal := strings.ToLower(filter.Value)

				if containsSearchColumns[filter.Column] {
					if !strings.Contains(cellValue, filterVal) {
						multiColumnMatch = false
						break
					}
				} else {
					if cellValue != filterVal {
						multiColumnMatch = false
						break
					}
				}
			} else if filter.Value != "" {
				multiColumnMatch = false
				break
			}
		}

		// Filtrado independiente por Rango de Edad
		ageMatch := true
		if hasMinAge || hasMaxAge {
			if edadColIndex != -1 && edadColIndex < len(row) {
				edadClean := regexp.MustCompile(`\D`).ReplaceAllString(row[edadColIndex], "")
				if edadNum, err := strconv.Atoi(edadClean); err == nil {
					if hasMinAge && edadNum < minAgeVal {
						ageMatch = false
					}
					if hasMaxAge && edadNum > maxAgeVal {
						ageMatch = false
					}
				} else {
					ageMatch = false
				}
			} else {
				ageMatch = false
			}
		}

		if globalMatch && multiColumnMatch && ageMatch {
			filtered = append(filtered, IndexedRow{Index: i + 1, Cells: row})
		}
	}

	data := make([]map[string]string, 0, length)
	for i := start; i < len(filtered) && len(data) < length; i++ {
		row := filtered[i]
		rec := map[string]string{}
		rec["__row"] = strconv.Itoa(row.Index)

		for j, key := range keys {
			val := ""
			if j < len(row.Cells) {
				val = row.Cells[j]
			}
			rec[strings.TrimSpace(key)] = val
		}
		data = append(data, rec)
	}

	resp := DTResponse{
		Draw:            draw,
		RecordsTotal:    len(rows) - 1,
		RecordsFiltered: len(filtered),
		Data:            data,
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(resp)
}

// updateExcelData maneja tanto la inserción/actualización de datos como la eliminación segura de registros excluidos.
func updateExcelData(w http.ResponseWriter, r *http.Request) {
	descargarDeDropbox()

	excelMutex.Lock()
	defer excelMutex.Unlock()

	fmt.Println("--- LOG: Endpoint /api/update-excel invocado en lote. ---")

	var req struct {
		Comunidad string              `json:"comunidad"`
		Torre     string              `json:"torre"`
		Casa      string              `json:"casa"`
		Datos     []map[string]string `json:"datos"`
	}
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		fmt.Printf("--- ERROR: No se pudo decodificar el payload JSON: %v ---\n", err)
		http.Error(w, "payload inválido", http.StatusBadRequest)
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	rows, _ := f.GetRows(PRIMERA_HOJA)
	headers := rows[0]

	// 1. Localizar los índices de columnas clave
	comIdx, torreIdx, casaIdx := -1, -1, -1
	for i, h := range headers {
		headerClean := strings.TrimSpace(strings.ToLower(h))
		if headerClean == "comunidad" {
			comIdx = i
		}
		if headerClean == "torre" {
			torreIdx = i
		}
		if headerClean == "numero de casa / apto" || headerClean == "casa" {
			casaIdx = i
		}
	}

	// 2. Identificar registros a eliminar de este núcleo familiar
	var rowsToDelete []int
	if req.Comunidad != "" && req.Torre != "" && req.Casa != "" && comIdx != -1 && torreIdx != -1 && casaIdx != -1 {
		existingRowsInExcel := make(map[int]bool)
		for i, row := range rows {
			if i == 0 {
				continue
			}
			if len(row) > comIdx && len(row) > torreIdx && len(row) > casaIdx {
				if row[comIdx] == req.Comunidad && row[torreIdx] == req.Torre && row[casaIdx] == req.Casa {
					existingRowsInExcel[i+1] = true // Formato base 1
				}
			}
		}

		payloadRows := make(map[int]bool)
		for _, fila := range req.Datos {
			rowNumStr := fila["__row"]
			if rowNum, err := strconv.Atoi(rowNumStr); err == nil {
				payloadRows[rowNum] = true
			}
		}

		for rNum := range existingRowsInExcel {
			if !payloadRows[rNum] {
				rowsToDelete = append(rowsToDelete, rNum)
			}
		}
	}

	// 3. Procesar las actualizaciones e inserciones
	nextAvailableRow := len(rows) + 1
	for _, fila := range req.Datos {
		rowNumStr := fila["__row"]
		rowNum, err := strconv.Atoi(rowNumStr)

		// Es un registro nuevo (sin ID de fila válido)
		if err != nil || strings.HasPrefix(rowNumStr, "new_") {
			fmt.Printf("--- LOG: Insertando nuevo miembro en la fila %d\n", nextAvailableRow)
			for colIndex, key := range headers {
				cleanKey := strings.TrimSpace(key)
				if val, ok := fila[cleanKey]; ok {
					cell, _ := excelize.CoordinatesToCellName(colIndex+1, nextAvailableRow)
					f.SetCellValue(PRIMERA_HOJA, cell, val)
				}
			}
			nextAvailableRow++
		} else {
			// Es una actualización, verificamos que no esté marcado para eliminación
			isDeleted := false
			for _, d := range rowsToDelete {
				if d == rowNum {
					isDeleted = true
					break
				}
			}
			if !isDeleted {
				fmt.Printf("--- LOG: Actualizando miembro en la fila %d\n", rowNum)
				for colIndex, key := range headers {
					cleanKey := strings.TrimSpace(key)
					if val, ok := fila[cleanKey]; ok {
						cell, _ := excelize.CoordinatesToCellName(colIndex+1, rowNum)
						f.SetCellValue(PRIMERA_HOJA, cell, val)
					}
				}
			}
		}
	}

	// 4. Ejecutar las eliminaciones acumuladas en orden descendente
	if len(rowsToDelete) > 0 {
		// Ordenamos de mayor a menor para evitar desalineaciones por corrimiento de índices
		for i := 0; i < len(rowsToDelete); i++ {
			for j := i + 1; j < len(rowsToDelete); j++ {
				if rowsToDelete[i] < rowsToDelete[j] {
					rowsToDelete[i], rowsToDelete[j] = rowsToDelete[j], rowsToDelete[i]
				}
			}
		}

		fmt.Printf("--- LOG: Removiendo filas del Excel en orden descendente: %v\n", rowsToDelete)
		for _, rNum := range rowsToDelete {
			if err := f.RemoveRow(PRIMERA_HOJA, rNum); err != nil {
				fmt.Printf("--- ERROR: Falló la eliminación de la fila %d: %v\n", rNum, err)
			}
		}
	}

	if err := f.Save(); err != nil {
		http.Error(w, "no se guardó el Excel", http.StatusInternalServerError)
		return
	}

	// Sincronización asíncrona con Dropbox
	go subirADropbox()

	fmt.Println("--- LOG: ¡Archivo Excel guardado y sincronizado exitosamente! ---")
	w.WriteHeader(http.StatusOK)

	addLog("Base de Datos: Se modificó la composición de un núcleo familiar")
}

// exportToExcel
// Exportar a Excel respetando filtros y rango de edad
// Exportar a Excel respetando filtros y rango de edad
func exportToExcel(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	search := strings.ToLower(r.URL.Query().Get("search[value]"))

	minAgeStr := r.URL.Query().Get("minAge")
	maxAgeStr := r.URL.Query().Get("maxAge")
	var minAgeVal, maxAgeVal int
	var hasMinAge, hasMaxAge bool

	if minAgeStr != "" {
		if val, err := strconv.Atoi(minAgeStr); err == nil {
			minAgeVal = val
			hasMinAge = true
		}
	}
	if maxAgeStr != "" {
		if val, err := strconv.Atoi(maxAgeStr); err == nil {
			maxAgeVal = val
			hasMaxAge = true
		}
	}

	filtersJSON := r.URL.Query().Get("filters")
	var activeFilters []Filter
	if filtersJSON != "" {
		if err := json.Unmarshal([]byte(filtersJSON), &activeFilters); err != nil {
			http.Error(w, "Filtros inválidos", http.StatusBadRequest)
			return
		}
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "no se pudo abrir el Excel", 500)
		return
	}

	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) == 0 {
		http.Error(w, "error leyendo filas", 500)
		return
	}

	headers := rows[0]

	edadColIndex := -1
	for j, key := range headers {
		if strings.TrimSpace(strings.ToLower(key)) == "edad" {
			edadColIndex = j
			break
		}
	}

	filteredRows := make([][]string, 0)
	for i := 1; i < len(rows); i++ {
		row := rows[i]

		globalMatch := false
		if search == "" {
			globalMatch = true
		} else {
			for _, cell := range row {
				if strings.Contains(strings.ToLower(cell), search) {
					globalMatch = true
					break
				}
			}
		}

		multiColumnMatch := true
		for _, filter := range activeFilters {
			if strings.TrimSpace(filter.Value) == "" {
				continue
			}

			filterColumnIndex := -1
			for j, key := range headers {
				if strings.TrimSpace(key) == filter.Column {
					filterColumnIndex = j
					break
				}
			}

			if filterColumnIndex != -1 && filterColumnIndex < len(row) {
				cellValue := strings.ToLower(strings.TrimSpace(row[filterColumnIndex]))
				filterVal := strings.ToLower(filter.Value)

				if containsSearchColumns[filter.Column] {
					if !strings.Contains(cellValue, filterVal) {
						multiColumnMatch = false
						break
					}
				} else {
					if cellValue != filterVal {
						multiColumnMatch = false
						break
					}
				}
			} else if filter.Value != "" {
				multiColumnMatch = false
				break
			}
		}

		ageMatch := true
		if hasMinAge || hasMaxAge {
			if edadColIndex != -1 && edadColIndex < len(row) {
				edadClean := regexp.MustCompile(`\D`).ReplaceAllString(row[edadColIndex], "")
				if edadNum, err := strconv.Atoi(edadClean); err == nil {
					if hasMinAge && edadNum < minAgeVal {
						ageMatch = false
					}
					if hasMaxAge && edadNum > maxAgeVal {
						ageMatch = false
					}
				} else {
					ageMatch = false
				}
			} else {
				ageMatch = false
			}
		}

		if globalMatch && multiColumnMatch && ageMatch {
			filteredRows = append(filteredRows, row)
		}
	}

	exportFile := excelize.NewFile()
	sheetName := "Reporte"
	index := exportFile.NewSheet(sheetName)
	exportFile.SetActiveSheet(index)

	for colIndex, header := range headers {
		cell := fmt.Sprintf("%s%d", columnLetter(colIndex), 1)
		exportFile.SetCellValue(sheetName, cell, header)
	}

	for rowIndex, rowData := range filteredRows {
		for colIndex, cellValue := range rowData {
			cell := fmt.Sprintf("%s%d", columnLetter(colIndex), rowIndex+2)
			exportFile.SetCellValue(sheetName, cell, cellValue)
		}
	}

	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", "attachment; filename=reporte_habitantes.xlsx")

	if err := exportFile.Write(w); err != nil {
		http.Error(w, "no se pudo escribir el archivo Excel", http.StatusInternalServerError)
	}
}

// exportToPDF
// Exportar a PDF respetando filtros y rango de edad
func exportToPDF(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	search := r.URL.Query().Get("search[value]")

	minAgeStr := r.URL.Query().Get("minAge")
	maxAgeStr := r.URL.Query().Get("maxAge")
	var minAgeVal, maxAgeVal int
	var hasMinAge, hasMaxAge bool

	if minAgeStr != "" {
		if val, err := strconv.Atoi(minAgeStr); err == nil {
			minAgeVal = val
			hasMinAge = true
		}
	}
	if maxAgeStr != "" {
		if val, err := strconv.Atoi(maxAgeStr); err == nil {
			maxAgeVal = val
			hasMaxAge = true
		}
	}

	filtersJSON := r.URL.Query().Get("filters")
	var activeFilters []Filter
	if filtersJSON != "" {
		if err := json.Unmarshal([]byte(filtersJSON), &activeFilters); err != nil {
			http.Error(w, "Filtros inválidos", http.StatusBadRequest)
			return
		}
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "no se pudo abrir el Excel", http.StatusInternalServerError)
		return
	}

	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) == 0 {
		http.Error(w, "error leyendo filas", http.StatusInternalServerError)
		return
	}

	allHeaders := rows[0]
	selectedColumns := []string{"Nombre completo", "Cedula de identidad", "Edad", "Genero"}

	var displayHeaders []string
	headerIndexMap := make(map[string]int)
	for i, header := range allHeaders {
		cleanHeader := strings.TrimSpace(header)
		for _, selectedCol := range selectedColumns {
			if cleanHeader == selectedCol {
				displayHeaders = append(displayHeaders, cleanHeader)
				headerIndexMap[cleanHeader] = i
				break
			}
		}
	}

	edadColIndex := -1
	for j, key := range allHeaders {
		if strings.TrimSpace(strings.ToLower(key)) == "edad" {
			edadColIndex = j
			break
		}
	}

	filteredData := make([]map[string]string, 0)
	for i := 1; i < len(rows); i++ {
		row := rows[i]

		globalMatch := false
		if search == "" {
			globalMatch = true
		} else {
			for _, cell := range row {
				if strings.Contains(strings.ToLower(cell), strings.ToLower(search)) {
					globalMatch = true
					break
				}
			}
		}

		multiColumnMatch := true
		for _, filter := range activeFilters {
			if strings.TrimSpace(filter.Value) == "" {
				continue
			}

			filterColumnIndex := -1
			for j, key := range allHeaders {
				if strings.TrimSpace(key) == filter.Column {
					filterColumnIndex = j
					break
				}
			}

			if filterColumnIndex != -1 && filterColumnIndex < len(row) {
				cellValue := strings.ToLower(strings.TrimSpace(row[filterColumnIndex]))
				filterVal := strings.ToLower(filter.Value)

				if containsSearchColumns[filter.Column] {
					if !strings.Contains(cellValue, filterVal) {
						multiColumnMatch = false
						break
					}
				} else {
					if cellValue != filterVal {
						multiColumnMatch = false
						break
					}
				}
			} else if filter.Value != "" {
				multiColumnMatch = false
				break
			}
		}

		ageMatch := true
		if hasMinAge || hasMaxAge {
			if edadColIndex != -1 && edadColIndex < len(row) {
				edadClean := regexp.MustCompile(`\D`).ReplaceAllString(row[edadColIndex], "")
				if edadNum, err := strconv.Atoi(edadClean); err == nil {
					if hasMinAge && edadNum < minAgeVal {
						ageMatch = false
					}
					if hasMaxAge && edadNum > maxAgeVal {
						ageMatch = false
					}
				} else {
					ageMatch = false
				}
			} else {
				ageMatch = false
			}
		}

		if globalMatch && multiColumnMatch && ageMatch {
			rowData := make(map[string]string)
			for _, header := range displayHeaders {
				originalIndex := headerIndexMap[header]
				val := ""
				if originalIndex < len(row) {
					val = row[originalIndex]
				}
				rowData[header] = val
			}
			filteredData = append(filteredData, rowData)
		}
	}

	data := struct {
		Headers       []string
		Rows          []map[string]string
		Search        string
		ActiveFilters []Filter
		RowCount      int
		MinAge        string
		MaxAge        string
	}{
		Headers:       displayHeaders,
		Rows:          filteredData,
		Search:        search,
		ActiveFilters: activeFilters,
		RowCount:      len(filteredData),
		MinAge:        minAgeStr,
		MaxAge:        maxAgeStr,
	}

	htmlTemplate := `
	<!DOCTYPE html>
	<html>
	<head>
		<meta charset="UTF-8">
		<title>Reporte de Habitantes</title>
		<style>
			body { font-family: Arial, sans-serif; margin: 20px; }
			h1 { text-align: center; color: #333; }
			h3 { text-align: center; color: #333; }
			p { color: #666; margin-bottom: 5px; }
			.filters-summary { margin-bottom: 20px; border: 1px solid #eee; padding: 10px; background-color: #f9f9f9; }
			.filters-summary p { margin: 0; }
			table { width: 100%; border-collapse: collapse; margin-top: 20px; }
			th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
			th { background-color: #f2f2f2; }
			tr:nth-child(even) { background-color: #f9f9f9; }
		</style>
	</head>
	<body>
		<h1>Reporte de Habitantes</h1>

		<div class="filters-summary">
			<h3>Filtros Aplicados:</h3>
			{{if .Search}}
			<p><strong>Búsqueda Global:</strong> "{{.Search}}"</p>
			{{end}}
			{{if .MinAge}}
			<p><strong>Edad Mínima:</strong> {{.MinAge}} años</p>
			{{end}}
			{{if .MaxAge}}
			<p><strong>Edad Máxima:</strong> {{.MaxAge}} años</p>
			{{end}}
			{{if .ActiveFilters}}
				{{range .ActiveFilters}}
				<p><strong>"{{.Column}}":</strong> "{{.Value}}"</p>
				{{end}}
			{{else if not .Search}}
			<p>No se aplicaron filtros específicos.</p>
			{{end}}
		</div>

		<h3>Cantidad de filas filtradas: {{.RowCount}}</h3>
		<table>
			<thead>
				<tr>
					{{range .Headers}}
					<th>{{.}}</th>
					{{end}}
				</tr>
			</thead>
			<tbody>
			{{range $rowIndex, $row := .Rows}}
				<tr>
				{{range $colIndex, $header := $.Headers}}
					<td>{{index $row $header}}</td>
				{{end}}
				</tr>
			{{end}}
			</tbody>
		</table>
	</body>
	</html>
	`

	tmpl, err := template.New("pdfReport").Parse(htmlTemplate)
	if err != nil {
		http.Error(w, "error al parsear el template HTML: "+err.Error(), http.StatusInternalServerError)
		return
	}

	var htmlBuffer bytes.Buffer
	if err := tmpl.Execute(&htmlBuffer, data); err != nil {
		http.Error(w, "error al ejecutar el template HTML: "+err.Error(), http.StatusInternalServerError)
		return
	}

	pdfg, err := wkhtml.NewPDFGenerator()
	if err != nil {
		http.Error(w, "no se pudo crear el generador de PDF: "+err.Error(), http.StatusInternalServerError)
		return
	}

	pdfg.AddPage(wkhtml.NewPageReader(bytes.NewReader(htmlBuffer.Bytes())))
	pdfg.PageSize.Set(wkhtml.PageSizeA4)
	pdfg.Orientation.Set(wkhtml.OrientationPortrait)

	err = pdfg.Create()
	if err != nil {
		http.Error(w, "no se pudo generar el PDF: "+err.Error(), http.StatusInternalServerError)
		return
	}

	w.Header().Set("Content-Type", "application/pdf")
	w.Header().Set("Content-Disposition", "attachment; filename=reporte_habitantes.pdf")
	w.Header().Set("Content-Length", strconv.Itoa(len(pdfg.Bytes())))

	if _, err := w.Write(pdfg.Bytes()); err != nil {
		http.Error(w, "no se pudo escribir el archivo PDF: "+err.Error(), http.StatusInternalServerError)
		return
	}
}

func exportVotantesPDF(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	search := strings.ToLower(r.URL.Query().Get("search[value]"))
	torreFiltro := strings.ToLower(r.URL.Query().Get("torre"))
	orden := r.URL.Query().Get("orden")

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "No se pudo abrir el Excel", http.StatusInternalServerError)
		return
	}

	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) < 2 {
		http.Error(w, "El archivo está vacío o no se leyó correctamente", http.StatusInternalServerError)
		return
	}

	headers := rows[0]

	comIdx, torreIdx, casaIdx, parentescoIdx, nombreIdx, cedulaIdx, edadIdx := -1, -1, -1, -1, -1, -1, -1

	for i, h := range headers {
		cleanH := strings.TrimSpace(strings.ToLower(h))
		switch cleanH {
		case "comunidad":
			comIdx = i
		case "torre":
			torreIdx = i
		case "numero de casa / apto", "casa":
			casaIdx = i
		case "parentesco":
			parentescoIdx = i
		case "nombre completo":
			nombreIdx = i
		case "cedula de identidad":
			cedulaIdx = i
		case "edad":
			edadIdx = i
		}
	}

	type Votante struct {
		Comunidad  string
		ConjRes    string
		CasaApto   string
		Parentesco string
		Nombre     string
		Cedula     string
	}

	var listaVotantes []Votante

	for i := 1; i < len(rows); i++ {
		row := rows[i]

		getVal := func(idx int) string {
			if idx != -1 && idx < len(row) {
				return strings.TrimSpace(row[idx])
			}
			return ""
		}

		cedula := cleanCedula(getVal(cedulaIdx))
		edadRaw := getVal(edadIdx)
		edadClean := regexp.MustCompile(`\D`).ReplaceAllString(edadRaw, "")
		edadNum, _ := strconv.Atoi(edadClean)

		com := getVal(comIdx)
		torre := getVal(torreIdx)
		casa := getVal(casaIdx)
		parentesco := getVal(parentescoIdx)
		nombre := getVal(nombreIdx)

		if edadNum < 16 || cedula == "" {
			continue
		}

		if torreFiltro != "" && strings.ToLower(torre) != torreFiltro {
			continue
		}

		if search != "" {
			matchNombre := strings.Contains(strings.ToLower(nombre), search)
			matchCedula := strings.Contains(strings.ToLower(cedula), search)
			matchUbicacion := strings.Contains(strings.ToLower(com+" "+torre+" "+casa), search)

			if !matchNombre && !matchCedula && !matchUbicacion {
				continue
			}
		}

		listaVotantes = append(listaVotantes, Votante{
			Comunidad:  com,
			ConjRes:    torre,
			CasaApto:   casa,
			Parentesco: parentesco,
			Nombre:     nombre,
			Cedula:     cedula,
		})
	}

	// Lógica de Ordenamiento
	sort.Slice(listaVotantes, func(i, j int) bool {
		if orden == "cedula" {
			// Orden numérico por cédula
			valI, _ := strconv.Atoi(regexp.MustCompile(`\D`).ReplaceAllString(listaVotantes[i].Cedula, ""))
			valJ, _ := strconv.Atoi(regexp.MustCompile(`\D`).ReplaceAllString(listaVotantes[j].Cedula, ""))
			return valI < valJ
		} else if orden == "nombre" {
			// Orden alfabético por nombre
			return strings.ToLower(listaVotantes[i].Nombre) < strings.ToLower(listaVotantes[j].Nombre)
		} else {
			// Orden por Jerarquía (Predeterminado: Comunidad -> Torre -> Casa)
			if listaVotantes[i].Comunidad != listaVotantes[j].Comunidad {
				return listaVotantes[i].Comunidad < listaVotantes[j].Comunidad
			}
			if listaVotantes[i].ConjRes != listaVotantes[j].ConjRes {
				// Intentar orden numérico de torre si es posible
				tI, errI := strconv.Atoi(regexp.MustCompile(`\D`).ReplaceAllString(listaVotantes[i].ConjRes, ""))
				tJ, errJ := strconv.Atoi(regexp.MustCompile(`\D`).ReplaceAllString(listaVotantes[j].ConjRes, ""))
				if errI == nil && errJ == nil {
					return tI < tJ
				}
				return listaVotantes[i].ConjRes < listaVotantes[j].ConjRes
			}
			// Por último, casa
			cI, errI := strconv.Atoi(regexp.MustCompile(`\D`).ReplaceAllString(listaVotantes[i].CasaApto, ""))
			cJ, errJ := strconv.Atoi(regexp.MustCompile(`\D`).ReplaceAllString(listaVotantes[j].CasaApto, ""))
			if errI == nil && errJ == nil {
				return cI < cJ
			}
			return listaVotantes[i].CasaApto < listaVotantes[j].CasaApto
		}
	})

	// Agrupamos los votantes en bloques/páginas de máximo 9 personas
	const registrosPorPagina = 11
	type Pagina struct {
		Filas []Votante
	}

	var paginas []Pagina
	for i := 0; i < len(listaVotantes); i += registrosPorPagina {
		end := i + registrosPorPagina
		if end > len(listaVotantes) {
			end = len(listaVotantes)
		}
		paginas = append(paginas, Pagina{
			Filas: listaVotantes[i:end],
		})
	}

	htmlTemplate := `
	<!DOCTYPE html>
	<html>
	<head>
		<meta charset="UTF-8">
		<title>Cuaderno electoral</title>
		<style>
			@page { 
				size: letter landscape; 
				margin: 8mm; 
			}
			body { 
				font-family: Arial, sans-serif; 
				margin: 0; 
				padding: 0; 
			}
			.page {
				page-break-after: always;
				height: 100%;
			}
			.page:last-child {
				page-break-after: auto;
			}
			table { 
				width: 100%; 
				border-collapse: collapse; 
				table-layout: fixed; 
			}
			th, td { 
				border: 1px solid #000; 
				padding: 4px 6px; 
				font-size: 13px; 
				vertical-align: middle; 
			}
			th { 
				background-color: #f2f2f2; 
				text-align: center; 
				font-weight: bold; 
				text-transform: uppercase; 
				height: 28px;
			}
			
			/* Altura fija de fila (~58px) para dar espacio cómodo a la huella dactilar y firma */
			tbody td {
				height: 70px;
			}

			.col-num { width: 32px; text-align: center; font-weight: bold; }
			.col-com { width: 95px; }
			.col-conj { width: 65px; text-align: center; }
			.col-casa { width: 75px; text-align: center; }
			.col-parentesco { width: 85px; }
			.col-nombre { width: 210px; }
			.col-cedula { width: 90px; text-align: center; }
			.col-firma { width: 110px; }
			.col-huella { width: 110px; }
		</style>
	</head>
	<body>
		{{range $pIdx, $p := .Paginas}}
		<div class="page">
			<table>
				<thead>
					<tr>
						<th class="col-num">N°</th>
						<th class="col-com">COMUNIDAD</th>
						<th class="col-conj">CONJ. RES.</th>
						<th class="col-casa">CASA O APTO. N°</th>
						<th class="col-parentesco">PARENTESCO</th>
						<th class="col-nombre">NOMBRE COMPLETO</th>
						<th class="col-cedula">DOCUMENTO IDENTIDAD</th>
						<th class="col-firma">FIRMA</th>
						<th class="col-huella">HUELLA</th>
					</tr>
				</thead>
				<tbody>
				{{range $fIdx, $v := $p.Filas}}
					<tr>
						<td class="col-num">{{calcularIndice $pIdx $fIdx}}</td>
						<td class="col-com">{{$v.Comunidad}}</td>
						<td class="col-conj">{{$v.ConjRes}}</td>
						<td class="col-casa">{{$v.CasaApto}}</td>
						<td class="col-parentesco">{{$v.Parentesco}}</td>
						<td class="col-nombre"><strong>{{$v.Nombre}}</strong></td>
						<td class="col-cedula">{{$v.Cedula}}</td>
						<td class="col-firma"></td>
						<td class="col-huella"></td>
					</tr>
				{{end}}
				</tbody>
			</table>
		</div>
		{{end}}
	</body>
	</html>`

	dataTemplate := struct {
		Paginas []Pagina
	}{
		Paginas: paginas,
	}

	tmpl, err := template.New("pdfReport").Funcs(template.FuncMap{
		"calcularIndice": func(pageIdx, filaIdx int) int {
			return (pageIdx * registrosPorPagina) + filaIdx + 1
		},
	}).Parse(htmlTemplate)

	if err != nil {
		http.Error(w, "Error al procesar plantilla: "+err.Error(), http.StatusInternalServerError)
		return
	}

	var htmlBuffer bytes.Buffer
	if err := tmpl.Execute(&htmlBuffer, dataTemplate); err != nil {
		http.Error(w, "Error al construir HTML: "+err.Error(), http.StatusInternalServerError)
		return
	}

	pdfg, err := wkhtml.NewPDFGenerator()
	if err != nil {
		http.Error(w, "Error al crear generador PDF: "+err.Error(), http.StatusInternalServerError)
		return
	}

	pdfg.AddPage(wkhtml.NewPageReader(bytes.NewReader(htmlBuffer.Bytes())))
	// Configuración explícita a Hoja Carta (Letter)
	pdfg.PageSize.Set(wkhtml.PageSizeLetter)
	pdfg.Orientation.Set(wkhtml.OrientationLandscape)

	if err := pdfg.Create(); err != nil {
		http.Error(w, "Error al renderizar el PDF: "+err.Error(), http.StatusInternalServerError)
		return
	}

	w.Header().Set("Content-Type", "application/pdf")
	w.Header().Set("Content-Disposition", "attachment; filename=cuaderno_votantes.pdf")
	w.Header().Set("Content-Length", strconv.Itoa(len(pdfg.Bytes())))
	w.Write(pdfg.Bytes())
}

// convierte 0 -> A, 25 -> Z, 26 -> AA, 27 -> AB, etc.
func columnLetter(idx int) string {
	var col string
	for idx >= 0 {
		col = string(rune('A'+(idx%26))) + col
		idx = idx/26 - 1
	}
	return col
}

// ------------------- INICIO DEL SERVIDOR -------------------------
// Estructura para las actividades del calendario
type Activity struct {
	ID          int    `json:"id"`
	Title       string `json:"title"`
	Description string `json:"description"`
	StartDate   string `json:"start_date"`
	EndDate     string `json:"end_date"`
	Time        string `json:"time"`
	Location    string `json:"location"`
	Image       string `json:"image"`
}

var activities = []Activity{
	{ID: 1, Title: "Reunión Consejo Comunal", StartDate: "2025-09-15", Time: "10:00", Location: "Salón Comunal", Image: ""},
	{ID: 2, Title: "Jornada de Vacunación", StartDate: "2025-09-20", EndDate: "2025-09-21", Description: "Jornada de vacunación para niños y adultos mayores.", Location: "Centro de Salud", Image: ""},
}
var lastActivityID = 2

// Carga las actividades desde el archivo JSON al iniciar
func loadActivitiesFromFile() {
	if _, err := os.Stat(ACTIVITIES_FILE); os.IsNotExist(err) {
		return
	}
	data, err := ioutil.ReadFile(ACTIVITIES_FILE)
	if err != nil {
		fmt.Println("Error al leer actividades:", err)
		return
	}
	json.Unmarshal(data, &activities)

	for _, a := range activities {
		if a.ID > lastActivityID {
			lastActivityID = a.ID
		}
	}
}

// Guarda el slice de actividades en el archivo JSON
func saveActivitiesToFile() {
	data, err := json.MarshalIndent(activities, "", "  ")
	if err != nil {
		fmt.Println("Error al codificar actividades:", err)
		return
	}
	err = ioutil.WriteFile(ACTIVITIES_FILE, data, 0644)
	if err != nil {
		fmt.Println("Error al guardar actividades:", err)
	}
}

func getActivitiesHandler(w http.ResponseWriter, r *http.Request) {
	events := make([]map[string]interface{}, 0, len(activities))
	for _, a := range activities {
		event := map[string]interface{}{
			"id":    a.ID,
			"title": a.Title,
			"start": a.StartDate,
			"end":   a.EndDate,
			"extendedProps": map[string]string{
				"description": a.Description,
				"time":        a.Time,
				"location":    a.Location,
				"image":       a.Image,
			},
		}
		if a.EndDate != "" {
			end, _ := time.Parse("2006-01-02", a.EndDate)
			end = end.Add(24 * time.Hour)
			event["end"] = end.Format("2006-01-02")
		}
		events = append(events, event)
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(events)
}

func addActivityHandler(w http.ResponseWriter, r *http.Request) {
	if err := r.ParseMultipartForm(10 << 20); err != nil {
		http.Error(w, "No se pudo parsear el formulario", http.StatusBadRequest)
		return
	}

	var newActivity Activity
	newActivity.Title = r.FormValue("title")
	newActivity.Description = r.FormValue("description")
	newActivity.StartDate = r.FormValue("start_date")
	newActivity.EndDate = r.FormValue("end_date")
	newActivity.Time = r.FormValue("time")
	newActivity.Location = r.FormValue("location")

	file, header, err := r.FormFile("image")
	if err == nil {
		defer file.Close()
		filename := fmt.Sprintf("%d-%s", time.Now().UnixNano(), filepath.Base(header.Filename))
		ext := strings.ToLower(filepath.Ext(filename))
		if ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".gif" {
			http.Error(w, "Formato de imagen no permitido", http.StatusBadRequest)
			return
		}

		dst, err := os.Create(filepath.Join(uploadDir, filename))
		if err != nil {
			http.Error(w, "Error al guardar la imagen", http.StatusInternalServerError)
			return
		}
		defer dst.Close()
		io.Copy(dst, file)
		newActivity.Image = filename
	} else if err != http.ErrMissingFile {
		http.Error(w, "Error al procesar el archivo subido", http.StatusInternalServerError)
		return
	}

	lastActivityID++
	newActivity.ID = lastActivityID
	activities = append(activities, newActivity)

	saveActivitiesToFile()
	addLog("Calendario: Se agregó la actividad " + newActivity.Title)

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]bool{"success": true})
}

func editActivityHandler(w http.ResponseWriter, r *http.Request) {
	idStr := strings.TrimPrefix(r.URL.Path, "/api/activities/edit/")
	id, err := strconv.Atoi(idStr)
	if err != nil {
		http.Error(w, "ID inválido", http.StatusBadRequest)
		return
	}

	if err := r.ParseMultipartForm(10 << 20); err != nil {
		http.Error(w, "No se pudo parsear el formulario", http.StatusBadRequest)
		return
	}

	var updatedActivity Activity
	updatedActivity.Title = r.FormValue("title")
	updatedActivity.Description = r.FormValue("description")
	updatedActivity.StartDate = r.FormValue("start_date")
	updatedActivity.EndDate = r.FormValue("end_date")
	updatedActivity.Time = r.FormValue("time")
	updatedActivity.Location = r.FormValue("location")

	for i, a := range activities {
		if a.ID == id {
			activities[i].Title = updatedActivity.Title
			activities[i].Description = updatedActivity.Description
			activities[i].StartDate = updatedActivity.StartDate
			activities[i].EndDate = updatedActivity.EndDate
			activities[i].Time = updatedActivity.Time
			activities[i].Location = updatedActivity.Location

			file, header, err := r.FormFile("image")
			if err == nil {
				defer file.Close()
				if activities[i].Image != "" {
					os.Remove(filepath.Join(uploadDir, activities[i].Image))
				}
				filename := fmt.Sprintf("%d-%s", time.Now().UnixNano(), filepath.Base(header.Filename))
				ext := strings.ToLower(filepath.Ext(filename))
				if ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".gif" {
					http.Error(w, "Formato de imagen no permitido", http.StatusBadRequest)
					return
				}

				dst, err := os.Create(filepath.Join(uploadDir, filename))
				if err != nil {
					http.Error(w, "Error al guardar la nueva imagen", http.StatusInternalServerError)
					return
				}
				defer dst.Close()
				io.Copy(dst, file)
				activities[i].Image = filename
			} else if err != http.ErrMissingFile {
				http.Error(w, "Error al procesar la nueva imagen", http.StatusInternalServerError)
				return
			}
			break
		}
	}
	saveActivitiesToFile()
	addLog("Calendario: Se editó una actividad")

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]bool{"success": true})
}

func deleteActivityHandler(w http.ResponseWriter, r *http.Request) {
	idStr := strings.TrimPrefix(r.URL.Path, "/api/activities/delete/")
	id, err := strconv.Atoi(idStr)
	if err != nil {
		http.Error(w, "ID inválido", http.StatusBadRequest)
		return
	}

	for i, a := range activities {
		if a.ID == id {
			if a.Image != "" {
				err := os.Remove(filepath.Join(uploadDir, a.Image))
				if err != nil {
					fmt.Printf("Advertencia: no se pudo eliminar la imagen %s: %v\n", a.Image, err)
				}
			}
			activities = append(activities[:i], activities[i+1:]...)
			break
		}
	}
	saveActivitiesToFile()
	addLog("Calendario: Se eliminó una actividad")

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]bool{"success": true})
}

// Handler para el Historial
func getHistoryHandler(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(historyLogs)
}

// Permite descargar el archivo físico real
func downloadFullExcelHandler(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	if _, err := os.Stat(EXCEL_FILE); os.IsNotExist(err) {
		http.Error(w, "Archivo no encontrado", http.StatusNotFound)
		return
	}
	addLog("Exportación: Se descargó el archivo Excel completo")
	w.Header().Set("Content-Disposition", "attachment; filename=\""+EXCEL_FILE+"\"")
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	http.ServeFile(w, r, EXCEL_FILE)
}

// Reemplaza el archivo local y lo sube a Dropbox
func uploadFullExcelHandler(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	if err := r.ParseMultipartForm(10 << 20); err != nil {
		http.Error(w, "Error en formulario", http.StatusBadRequest)
		return
	}

	file, _, err := r.FormFile("excelFile")
	if err != nil {
		http.Error(w, "Archivo no encontrado", http.StatusBadRequest)
		return
	}
	defer file.Close()

	dst, err := os.Create(EXCEL_FILE)
	if err != nil {
		http.Error(w, "Error al crear archivo local", 500)
		return
	}
	defer dst.Close()

	if _, err := io.Copy(dst, file); err != nil {
		http.Error(w, "Error al guardar contenido", 500)
		return
	}

	// Sincronizar con Dropbox inmediatamente
	go subirADropbox()

	addLog("Importación: Se reemplazó la base de datos completa y se subió a Dropbox")
	w.WriteHeader(http.StatusOK)
}

func main() {
	os.MkdirAll(uploadDir, os.ModePerm)

	// CARGAR HISTORIAL PERSISTENTE
	loadLogsFromFile()
	loadActivitiesFromFile()

	fmt.Println("--- INICIALIZACIÓN: Verificando base de datos con la nube... ---")
	localMissingOrEmpty := false
	if info, err := os.Stat(EXCEL_FILE); os.IsNotExist(err) || info.Size() == 0 {
		localMissingOrEmpty = true
	}

	if localMissingOrEmpty {
		fmt.Println("Base de datos local faltante o vacía. Descargando última versión de Dropbox de manera síncrona...")
		err := descargarDeDropbox()
		if err != nil {
			fmt.Println("Error crítico en arranque: No se pudo descargar la base de datos inicial:", err)
		} else {
			fmt.Println("✅ Base de datos descargada correctamente en el arranque.")
		}
	} else {
		// El archivo ya existe localmente; se actualiza en segundo plano sin bloquear el arranque del sistema
		go func() {
			fmt.Println("Base de datos local detectada. Comprobando actualizaciones en segundo plano...")
			err := descargarDeDropbox()
			if err != nil {
				fmt.Println(" Sincronización en segundo plano omitida (modo sin conexión):", err)
			} else {
				fmt.Println(" Base de datos sincronizada con la última versión en la nube.")
			}
		}()
	}

	//  Rutas api
	http.HandleFunc("/api/activities", getActivitiesHandler)
	http.HandleFunc("/api/activities/add", addActivityHandler)
	http.HandleFunc("/api/activities/edit/", editActivityHandler)
	http.HandleFunc("/api/activities/delete/", deleteActivityHandler)
	http.HandleFunc("/api/delete-row", deleteRowHandler)
	http.HandleFunc("/importar", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/importar.html")
	})
	http.HandleFunc("/api/history", getHistoryHandler)
	http.HandleFunc("/api/bulk-import", bulkImportHandler)
	http.HandleFunc("/api/check-cedulas", checkCedulasHandler)
	http.HandleFunc("/api/get-person-by-cedula", getPersonByCedulaHandler)
	http.HandleFunc("/galeria", galleryHandler)
	http.HandleFunc("/upload", uploadHandler)
	http.HandleFunc("/delete/", deleteHandler)
	http.HandleFunc("/api/tree-data", getTreeData)
	http.HandleFunc("/api/get-people", getPeopleInHouse)
	http.HandleFunc("/api/pdf/export", exportToPDF)
	http.HandleFunc("/api/excel/export", exportToExcel)
	http.HandleFunc("/api/update-excel", updateExcelData)
	http.HandleFunc("/api/excel/columns", getColumns)
	http.HandleFunc("/api/excel", getData)
	http.HandleFunc("/api/excel/download-full", downloadFullExcelHandler)
	http.HandleFunc("/api/excel/upload-full", uploadFullExcelHandler)
	http.HandleFunc("/api/tree/rename", treeRenameHandler)
	http.HandleFunc("/api/tree/delete", treeDeleteHandler)
	http.HandleFunc("/api/tree/create", treeCreateHandler)
	http.HandleFunc("/api/search-people", searchPeopleHandler)
	http.HandleFunc("/api/tree/delete-multiple", treeDeleteMultipleHandler)
	// Registro de endpoints
	http.HandleFunc("/api/sync-status", syncStatusHandler)
	http.HandleFunc("/api/sync-force", forceSyncHandler)

	http.HandleFunc("/api/votantes/pdf", exportVotantesPDF)
	http.HandleFunc("/api/preview-import", previewImportHandler)

	// Inicializar log
	addLog("Sistema iniciado con sincronización Dropbox")

	http.HandleFunc("/editar-hogar", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/editar_hogar.html")
	})
	http.HandleFunc("/api/get-household-details", getHouseholdDetails)
	http.HandleFunc("/api/add-household", addHouseholdData)
	http.HandleFunc("/agregar-hogar", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/editar_hogar.html")
	})

	http.HandleFunc("/historia", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/historia.html")
	})

	http.HandleFunc("/listado_votantes", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/listado_votantes.html")
	})

	fs := http.FileServer(http.Dir("assets"))
	http.Handle("/assets/", http.StripPrefix("/assets/", fs))

	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/index.html")
	})
	http.HandleFunc("/base_de_datos", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/Base_de_Datos.html")
	})
	http.HandleFunc("/calendario", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/calendario.html")
	})
	http.HandleFunc("/comunidades", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/comunidades.html")
	})
	http.HandleFunc("/dashboard", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/dashboard.html")
	})

	http.HandleFunc("/api/egresos-tree-data", getEgresosTreeData)
	http.HandleFunc("/api/dashboard-stats", getDashboardStatsHandler)

	http.HandleFunc("/filtros", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/filtros.html")
	})

	// O si sirves las plantillas HTML manualmente:
	http.HandleFunc("/egresos", func(w http.ResponseWriter, r *http.Request) {
		http.ServeFile(w, r, "paginas/egresos.html")
	})

	go func() {
		fmt.Println("Servidor corriendo en http://localhost:8080")
		if err := http.ListenAndServe("127.0.0.1:8080", nil); err != nil {
			fmt.Println("Error:", err)
		}
	}()

	time.Sleep(500 * time.Millisecond)
	openBrowser("http://localhost:8080")
	select {}
}

//------------------- MANEJO DE LA GALERIA -------------------------

func galleryHandler(w http.ResponseWriter, r *http.Request) {
	files, err := ioutil.ReadDir(uploadDir)
	if err != nil {
		http.Error(w, "Error al leer imágenes", http.StatusInternalServerError)
		return
	}

	var images []ImageInfo
	for _, file := range files {
		if !file.IsDir() {
			images = append(images, ImageInfo{
				Filename:   file.Name(),
				UploadDate: file.ModTime().Format("02/01/2006 03:04 PM"),
			})
		}
	}

	tmpl, err := template.ParseFiles("paginas/galeria.html")
	if err != nil {
		http.Error(w, "Error al cargar plantilla: "+err.Error(), 500)
		return
	}

	data := GalleryData{
		Images:   images,
		Messages: []string{},
	}
	tmpl.Execute(w, data)
}

func uploadHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Redirect(w, r, "/galeria", http.StatusSeeOther)
		return
	}

	file, header, err := r.FormFile("file")
	if err != nil {
		http.Error(w, "Error al subir archivo", http.StatusBadRequest)
		return
	}
	defer file.Close()

	filename := filepath.Base(header.Filename)
	ext := strings.ToLower(filepath.Ext(filename))
	if ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".gif" {
		http.Error(w, "Formato no permitido", http.StatusBadRequest)
		return
	}

	dst, err := os.Create(filepath.Join(uploadDir, filename))
	if err != nil {
		http.Error(w, "Error al guardar imagen", http.StatusInternalServerError)
		return
	}
	defer dst.Close()

	io.Copy(dst, file)
	http.Redirect(w, r, "/galeria", http.StatusSeeOther)

	addLog("Galería: Nueva imagen subida")
}

func deleteHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Redirect(w, r, "/galeria", http.StatusSeeOther)
		return
	}

	filename := strings.TrimPrefix(r.URL.Path, "/delete/")
	filepath := filepath.Join(uploadDir, filename)

	if err := os.Remove(filepath); err != nil {
		http.Error(w, "No se pudo eliminar la imagen", http.StatusInternalServerError)
		return
	}

	http.Redirect(w, r, "/galeria", http.StatusSeeOther)

	addLog("Galería: Imagen eliminada")
}

func cleanExcelValue(val string) string {
	val = strings.TrimSpace(val)
	return strings.TrimSuffix(val, ".0")
}

// Helper para normalizar cédulas ignorando puntos y comas
func cleanCedula(cedula string) string {
	cedula = strings.ReplaceAll(cedula, ",", "")
	cedula = strings.ReplaceAll(cedula, ".", "")
	return strings.TrimSpace(cedula)
}

// getTreeData
func getTreeData(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		fmt.Println("Error al abrir el archivo Excel:", err)
		http.Error(w, "no se pudo abrir el Excel: "+err.Error(), http.StatusInternalServerError)
		return
	}

	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) < 2 {
		fmt.Println("Error: La hoja está vacía o no se pudo leer.")
		http.Error(w, "sheet vacío o no existe", http.StatusInternalServerError)
		return
	}

	headers := rows[0]
	comunidadIdx, torreIdx, casaIdx := -1, -1, -1
	for i, h := range headers {
		headerClean := strings.TrimSpace(strings.ToLower(h))
		switch headerClean {
		case "comunidad":
			comunidadIdx = i
		case "torre":
			torreIdx = i
		case "numero de casa / apto":
			casaIdx = i
		}
	}

	fmt.Println("--- Depurando la carga del árbol ---")
	fmt.Printf("Índice encontrado para 'Comunidad': %d\n", comunidadIdx)
	fmt.Printf("Índice encontrado para 'Torre': %d\n", torreIdx)
	fmt.Printf("Índice encontrado para 'Casa/Apto': %d\n", casaIdx)
	fmt.Println("------------------------------------")

	if comunidadIdx == -1 || torreIdx == -1 || casaIdx == -1 {
		errorMsg := "No se encontraron todas las columnas requeridas. Revisa que tu Excel tenga cabeceras llamadas 'COMUNIDAD', 'TORRE' y 'CASA O APTO'."
		fmt.Println(errorMsg)
		http.Error(w, errorMsg, http.StatusInternalServerError)
		return
	}

	tree := make(map[string]map[string]map[string]struct{})
	for _, row := range rows[1:] {
		if len(row) <= comunidadIdx || len(row) <= torreIdx || len(row) <= casaIdx {
			continue
		}
		comunidad := cleanExcelValue(row[comunidadIdx])
		torre := cleanExcelValue(row[torreIdx])
		casa := cleanExcelValue(row[casaIdx])
		if comunidad == "" || torre == "" || casa == "" {
			continue
		}
		if _, ok := tree[comunidad]; !ok {
			tree[comunidad] = make(map[string]map[string]struct{})
		}
		if _, ok := tree[comunidad][torre]; !ok {
			tree[comunidad][torre] = make(map[string]struct{})
		}
		tree[comunidad][torre][casa] = struct{}{}
	}

	var result []*TreeNode
	for comName, torres := range tree {
		comNode := &TreeNode{Text: comName, Type: "comunidad"}
		comNode.State.Opened = false

		for torreName, casas := range torres {
			torreNode := &TreeNode{Text: "Torre " + torreName, Type: "torre"}
			torreNode.State.Opened = false

			for casaName := range casas {
				casaNode := &TreeNode{Text: "Casa/Apto " + casaName, Type: "casa"}
				torreNode.Children = append(torreNode.Children, casaNode)
			}
			comNode.Children = append(comNode.Children, torreNode)
		}
		result = append(result, comNode)
	}

	fmt.Printf("Se encontraron %d comunidades para el árbol.\n", len(result))

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(result)
}

// getPeopleInHouse
func getPeopleInHouse(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	comunidad := r.URL.Query().Get("comunidad")
	torre := r.URL.Query().Get("torre")
	casa := r.URL.Query().Get("casa")

	if comunidad == "" || torre == "" || casa == "" {
		http.Error(w, "Faltan parámetros: comunidad, torre y casa son requeridos", http.StatusBadRequest)
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "no se pudo abrir el Excel", http.StatusInternalServerError)
		return
	}

	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) < 2 {
		http.Error(w, "sheet vacío o no existe", http.StatusInternalServerError)
		return
	}

	headers := rows[0]
	comIdx, torreIdx, casaIdx, parentescoIdx, nombresIdx, docIdx := -1, -1, -1, -1, -1, -1
	for i, h := range headers {
		headerClean := strings.TrimSpace(strings.ToLower(h))
		switch headerClean {
		case "comunidad":
			comIdx = i
		case "torre":
			torreIdx = i
		case "numero de casa / apto":
			casaIdx = i
		case "parentesco":
			parentescoIdx = i
		case "nombre completo":
			nombresIdx = i
		case "cedula de identidad":
			docIdx = i
		}
	}

	if comIdx == -1 || torreIdx == -1 || casaIdx == -1 || parentescoIdx == -1 || nombresIdx == -1 || docIdx == -1 {
		http.Error(w, "No se encontraron todas las columnas requeridas en el Excel", http.StatusInternalServerError)
		return
	}

	var people []Person
	for _, row := range rows[1:] {
		if len(row) > comIdx && len(row) > torreIdx && len(row) > casaIdx {
			if row[comIdx] == comunidad && row[torreIdx] == torre && row[casaIdx] == casa {
				person := Person{}
				if parentescoIdx < len(row) {
					person.Parentesco = row[parentescoIdx]
				}
				if nombresIdx < len(row) {
					person.Nombres = row[nombresIdx]
				}
				if docIdx < len(row) {
					person.Documento = row[docIdx]
				}
				people = append(people, person)
			}
		}
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(people)
}

// getHouseholdDetails
func getHouseholdDetails(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	comunidad := r.URL.Query().Get("comunidad")
	torre := r.URL.Query().Get("torre")
	casa := r.URL.Query().Get("casa")

	if comunidad == "" || torre == "" || casa == "" {
		http.Error(w, "Faltan parámetros", http.StatusBadRequest)
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "No se pudo abrir el Excel", http.StatusInternalServerError)
		return
	}

	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) < 2 {
		http.Error(w, "Sheet vacío o no existe", http.StatusInternalServerError)
		return
	}

	headers := rows[0]
	comIdx, torreIdx, casaIdx := -1, -1, -1
	for i, h := range headers {
		headerClean := strings.TrimSpace(strings.ToLower(h))
		if headerClean == "comunidad" {
			comIdx = i
		}
		if headerClean == "torre" {
			torreIdx = i
		}
		if headerClean == "numero de casa / apto" || headerClean == "casa" {
			casaIdx = i
		}
	}

	if comIdx == -1 || torreIdx == -1 || casaIdx == -1 {
		http.Error(w, "Columnas clave no encontradas", http.StatusInternalServerError)
		return
	}

	var householdData []map[string]string
	for i, row := range rows {
		if i == 0 {
			continue
		}

		if len(row) > comIdx && len(row) > torreIdx && len(row) > casaIdx {
			if row[comIdx] == comunidad && row[torreIdx] == torre && row[casaIdx] == casa {
				personData := make(map[string]string)
				personData["__row"] = strconv.Itoa(i + 1)
				for j, header := range headers {
					cleanHeader := strings.TrimSpace(header)
					if j < len(row) {
						personData[cleanHeader] = row[j]
					} else {
						personData[cleanHeader] = ""
					}
				}
				householdData = append(householdData, personData)
			}
		}
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(householdData)
}

// addHouseholdData
func addHouseholdData(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	var req struct {
		Datos []map[string]string `json:"datos"`
	}
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, "Payload inválido", http.StatusBadRequest)
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	rows, _ := f.GetRows(PRIMERA_HOJA)
	headers := rows[0]
	nextRow := len(rows) + 1

	headerMap := make(map[string]int)
	for i, h := range headers {
		headerMap[strings.TrimSpace(h)] = i
	}

	for _, persona := range req.Datos {
		for key, val := range persona {
			if colIndex, ok := headerMap[key]; ok {
				cell, _ := excelize.CoordinatesToCellName(colIndex+1, nextRow)
				f.SetCellValue(PRIMERA_HOJA, cell, val)
			}
		}
		nextRow++
	}

	if err := f.Save(); err != nil {
		http.Error(w, "No se guardó el Excel", http.StatusInternalServerError)
		return
	}
	w.WriteHeader(http.StatusOK)
}

// Descarga el archivo de Dropbox a la carpeta local (con bloqueo seguro)
func descargarDeDropbox() error {
	excelMutex.Lock()
	defer excelMutex.Unlock()
	return descargarDeDropboxInternal()
}

// descargarDeDropboxInternal realiza la descarga directa sin bloquear (para evitar deadlocks internos)
func descargarDeDropboxInternal() error {
	fmt.Println("--- DROPBOX: Descargando última versión del Excel... ---")

	token, err := obtenerAccessToken()
	if err != nil {
		return fmt.Errorf("error al obtener token: %v", err)
	}

	url := "https://content.dropboxapi.com/2/files/download"
	req, err := http.NewRequest("POST", url, nil)
	if err != nil {
		return err
	}

	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Dropbox-API-Arg", `{"path": "`+DROPBOX_FILE_PATH+`"}`)

	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("error al descargar: %s", resp.Status)
	}

	data, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return err
	}

	// Almacena el 'rev' retornado en los encabezados para usarlo posteriormente en el control de concurrencia
	respHeader := resp.Header.Get("Dropbox-API-Result")
	if respHeader != "" {
		var meta DropboxMetadata
		if err := json.Unmarshal([]byte(respHeader), &meta); err == nil {
			ultimoRevGuardado = meta.Rev
			fmt.Printf("--- DROPBOX: 'rev' capturado en descarga: %s ---\n", ultimoRevGuardado)
		}
	}

	fmt.Println("--- DROPBOX: Excel descargado ---")
	return ioutil.WriteFile(EXCEL_FILE, data, 0644)
}

// Sube el archivo local a Dropbox (con bloqueo seguro y manejo de conflictos)
func subirADropbox() error {
	excelMutex.Lock()
	defer excelMutex.Unlock()
	return subirADropboxInternal()
}

// subirADropboxInternal implementa el envío de archivos y el tratamiento del error de colisión (Conflict)
func subirADropboxInternal() error {
	fmt.Println("--- DROPBOX: Subiendo cambios a la nube... ---")

	token, err := obtenerAccessToken()
	if err != nil {
		return fmt.Errorf("error al obtener token: %v", err)
	}

	contenido, err := ioutil.ReadFile(EXCEL_FILE)
	if err != nil {
		return err
	}

	url := "https://content.dropboxapi.com/2/files/upload"
	req, err := http.NewRequest("POST", url, bytes.NewBuffer(contenido))
	if err != nil {
		return err
	}

	req.Header.Set("Authorization", "Bearer "+token)

	// Selección del modo: Se utiliza 'update' condicionado al rev almacenado, evitando sobrescribir cambios externos sin fusionar.
	var mode interface{}
	if ultimoRevGuardado != "" {
		mode = map[string]interface{}{
			".tag":   "update",
			"update": ultimoRevGuardado,
		}
	} else {
		mode = "overwrite"
	}

	apiArg := map[string]interface{}{
		"path":       DROPBOX_FILE_PATH,
		"mode":       mode,
		"autorename": false,
		"mute":       false,
	}
	argJSON, _ := json.Marshal(apiArg)

	req.Header.Set("Dropbox-API-Arg", string(argJSON))
	req.Header.Set("Content-Type", "application/octet-stream")

	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	// Detección de colisión concurrente (Conflicto / Rev viejo)
	if resp.StatusCode == http.StatusConflict || resp.StatusCode == http.StatusBadRequest {
		bodyBytes, _ := ioutil.ReadAll(resp.Body)
		bodyStr := string(bodyBytes)
		if strings.Contains(bodyStr, "conflict") || strings.Contains(bodyStr, "path/conflict") {
			fmt.Println("⚠️ CONFLICTO DETECTADO: El archivo en la nube fue modificado por otro usuario. Iniciando fusión de datos...")

			// 1. Guardar copia del archivo local con cambios antes de pisarlo con la descarga de la nube
			localBackupPath := EXCEL_FILE + ".local"
			err = ioutil.WriteFile(localBackupPath, contenido, 0644)
			if err != nil {
				return fmt.Errorf("error al guardar respaldo local: %v", err)
			}
			defer os.Remove(localBackupPath)

			// 2. Descargar la nueva versión de la nube (actualiza EXCEL_FILE y recupera el último rev)
			err = descargarDeDropboxInternal()
			if err != nil {
				return fmt.Errorf("error al descargar censo de la nube durante fusión: %v", err)
			}

			// 3. Fusionar cambios desde el archivo local de respaldo hacia el nuevo archivo descargado
			err = fusionarCambiosLocales(localBackupPath, EXCEL_FILE)
			if err != nil {
				return fmt.Errorf("error al fusionar cambios locales: %v", err)
			}

			// 4. Leer contenido fusionado definitivo
			contenidoFusionado, err := ioutil.ReadFile(EXCEL_FILE)
			if err != nil {
				return fmt.Errorf("error al leer archivo fusionado: %v", err)
			}

			// 5. Reintentar subida con el nuevo contenido fusionado y el nuevo rev obtenido
			return reintentarSubidaConNuevoRev(token, contenidoFusionado)
		}
		return fmt.Errorf("error al subir, status: %s, body: %s", resp.Status, bodyStr)
	}

	if resp.StatusCode != http.StatusOK {
		bodyBytes, _ := ioutil.ReadAll(resp.Body)
		return fmt.Errorf("error al subir: %s - %s", resp.Status, string(bodyBytes))
	}

	// Almacena el nuevo hash de revisión si la subida fue exitosa
	respHeader := resp.Header.Get("Dropbox-API-Result")
	if respHeader != "" {
		var meta DropboxMetadata
		if err := json.Unmarshal([]byte(respHeader), &meta); err == nil {
			ultimoRevGuardado = meta.Rev
			fmt.Printf("--- DROPBOX: Nuevo 'rev' tras subida exitosa: %s ---\n", ultimoRevGuardado)
		}
	}

	fmt.Println("--- DROPBOX: Cambios subidos a Excel satisfactoriamente ---")
	return nil
}

// reintentarSubidaConNuevoRev realiza un reintento de subida posterior a la fusión de datos
func reintentarSubidaConNuevoRev(token string, contenido []byte) error {
	fmt.Printf("--- DROPBOX: Reintentando subida con nuevo rev: %s ---\n", ultimoRevGuardado)

	url := "https://content.dropboxapi.com/2/files/upload"
	req, err := http.NewRequest("POST", url, bytes.NewBuffer(contenido))
	if err != nil {
		return err
	}

	req.Header.Set("Authorization", "Bearer "+token)

	apiArg := map[string]interface{}{
		"path":       DROPBOX_FILE_PATH,
		"mode":       map[string]interface{}{".tag": "update", "update": ultimoRevGuardado},
		"autorename": false,
		"mute":       false,
	}
	argJSON, _ := json.Marshal(apiArg)

	req.Header.Set("Dropbox-API-Arg", string(argJSON))
	req.Header.Set("Content-Type", "application/octet-stream")

	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		bodyBytes, _ := ioutil.ReadAll(resp.Body)
		return fmt.Errorf("error en reintento de subida: %s - %s", resp.Status, string(bodyBytes))
	}

	respHeader := resp.Header.Get("Dropbox-API-Result")
	if respHeader != "" {
		var meta DropboxMetadata
		if err := json.Unmarshal([]byte(respHeader), &meta); err == nil {
			ultimoRevGuardado = meta.Rev
			fmt.Printf("--- DROPBOX: Reintento exitoso. Nuevo 'rev': %s ---\n", ultimoRevGuardado)
		}
	}

	addLog("Base de Datos: Conflicto de sincronización resuelto y fusionado con éxito")
	return nil
}

// fusionarCambiosLocales realiza la mezcla de los registros locales en el archivo descargado de la nube (usando la cédula como ID único)
func fusionarCambiosLocales(localPath, cloudPath string) error {
	fLocal, err := excelize.OpenFile(localPath)
	if err != nil {
		return fmt.Errorf("error al abrir archivo local temporal: %v", err)
	}

	fCloud, err := excelize.OpenFile(cloudPath)
	if err != nil {
		return fmt.Errorf("error al abrir archivo descargado de la nube: %v", err)
	}

	rowsLocal, err := fLocal.GetRows(PRIMERA_HOJA)
	if err != nil {
		return fmt.Errorf("error al obtener filas locales: %v", err)
	}

	rowsCloud, err := fCloud.GetRows(PRIMERA_HOJA)
	if err != nil {
		return fmt.Errorf("error al obtener filas de la nube: %v", err)
	}

	if len(rowsLocal) < 2 {
		return nil // No hay registros locales para procesar
	}

	// Localización de la columna clave "Cedula de identidad"
	headers := rowsLocal[0]
	cedulaColIdx := -1
	cedulaHeaderNormalized := normalizeHeader("Cedula de identidad")
	for i, h := range headers {
		if normalizeHeader(h) == cedulaHeaderNormalized {
			cedulaColIdx = i
			break
		}
	}

	if cedulaColIdx == -1 {
		return fmt.Errorf("no se encontró la columna de cédula en el archivo para la fusión")
	}

	// Mapeo de filas existentes en el archivo de la nube por Cédula (índice base 1 compatible con excelize)
	cloudCedulaMap := make(map[string]int)
	for idx, row := range rowsCloud {
		if idx == 0 {
			continue
		}
		if len(row) > cedulaColIdx {
			cleanCedString := cleanCedula(row[cedulaColIdx])
			if cleanCedString != "" {
				cloudCedulaMap[cleanCedString] = idx + 1
			}
		}
	}

	nextAvailableCloudRow := len(rowsCloud) + 1
	for idx, rowLocal := range rowsLocal {
		if idx == 0 {
			continue
		}

		if len(rowLocal) <= cedulaColIdx {
			continue
		}

		cedLocal := cleanCedula(rowLocal[cedulaColIdx])
		if cedLocal == "" {
			continue
		}

		if cloudRowIdx, exists := cloudCedulaMap[cedLocal]; exists {
			// El registro ya existe en el archivo de la nube: Se actualiza con los valores locales más recientes
			for colIdx, valLocal := range rowLocal {
				cell, _ := excelize.CoordinatesToCellName(colIdx+1, cloudRowIdx)
				fCloud.SetCellValue(PRIMERA_HOJA, cell, valLocal)
			}
		} else {
			// El registro es nuevo: Se anexa al final de la hoja descargada
			for colIdx, valLocal := range rowLocal {
				cell, _ := excelize.CoordinatesToCellName(colIdx+1, nextAvailableCloudRow)
				fCloud.SetCellValue(PRIMERA_HOJA, cell, valLocal)
			}
			nextAvailableCloudRow++
		}
	}

	// Se guardan los cambios de forma directa en cloudPath (el cual corresponde a EXCEL_FILE)
	if err := fCloud.SaveAs(cloudPath); err != nil {
		return fmt.Errorf("error al guardar archivo fusionado: %v", err)
	}

	fmt.Println("--- FUSION: Fusión de cambios locales completada con éxito. ---")
	return nil
}

// Estructura para peticiones de edición macro de la jerarquía
type TreeEditRequest struct {
	Level     string `json:"level"` // "comunidad", "torre", "casa"
	Comunidad string `json:"comunidad"`
	Torre     string `json:"torre"`
	Casa      string `json:"casa"`
	NewName   string `json:"new_name"`
}

// Función auxiliar para obtener índices de columnas de manera tolerante a variaciones de nombre
func getColumnIndices(headers []string) (comIdx, torreIdx, casaIdx, cedulaIdx, nombreIdx int) {
	comIdx, torreIdx, casaIdx, cedulaIdx, nombreIdx = -1, -1, -1, -1, -1
	for i, h := range headers {
		clean := strings.TrimSpace(strings.ToLower(h))
		switch clean {
		case "comunidad":
			comIdx = i
		case "torre":
			torreIdx = i
		case "numero de casa / apto", "numero de casa o apto", "casa o apto", "casa / apto", "casa":
			casaIdx = i
		case "cedula de identidad", "cedula":
			cedulaIdx = i
		case "nombre completo", "nombres", "nombre":
			nombreIdx = i
		}
	}
	return
}

// treeRenameHandler modifica el nombre de una comunidad, torre o casa en todas las filas coincidentes
func treeRenameHandler(w http.ResponseWriter, r *http.Request) {
	descargarDeDropbox()
	excelMutex.Lock()
	defer excelMutex.Unlock()

	var req TreeEditRequest
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, "JSON inválido", http.StatusBadRequest)
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "No se pudo abrir el archivo Excel", http.StatusInternalServerError)
		return
	}

	rows, _ := f.GetRows(PRIMERA_HOJA)
	if len(rows) < 2 {
		http.Error(w, "El archivo Excel no contiene datos", http.StatusBadRequest)
		return
	}

	comIdx, torreIdx, casaIdx, _, _ := getColumnIndices(rows[0])
	if comIdx == -1 || torreIdx == -1 || casaIdx == -1 {
		http.Error(w, "Columnas de jerarquía requeridas no encontradas", http.StatusInternalServerError)
		return
	}

	updatedCount := 0
	for i := 1; i < len(rows); i++ {
		row := rows[i]
		if len(row) <= comIdx || len(row) <= torreIdx || len(row) <= casaIdx {
			continue
		}

		match := false
		colToUpdate := -1

		// Normalizamos la comparación para mitigar discrepancias menores
		dbCom := strings.TrimSpace(row[comIdx])
		dbTorre := strings.TrimSpace(row[torreIdx])
		dbCasa := strings.TrimSpace(row[casaIdx])

		if req.Level == "comunidad" && dbCom == strings.TrimSpace(req.Comunidad) {
			match = true
			colToUpdate = comIdx
		} else if req.Level == "torre" && dbCom == strings.TrimSpace(req.Comunidad) && dbTorre == strings.TrimSpace(req.Torre) {
			match = true
			colToUpdate = torreIdx
		} else if req.Level == "casa" && dbCom == strings.TrimSpace(req.Comunidad) && dbTorre == strings.TrimSpace(req.Torre) && dbCasa == strings.TrimSpace(req.Casa) {
			match = true
			colToUpdate = casaIdx
		}

		if match && colToUpdate != -1 {
			cell, _ := excelize.CoordinatesToCellName(colToUpdate+1, i+1)
			f.SetCellValue(PRIMERA_HOJA, cell, strings.TrimSpace(req.NewName))
			updatedCount++
		}
	}

	if err := f.Save(); err != nil {
		http.Error(w, "Error al guardar el archivo", http.StatusInternalServerError)
		return
	}

	go subirADropbox()
	addLog(fmt.Sprintf("Jerarquía: Se renombró %s '%s' a '%s'", req.Level, req.Comunidad, req.NewName))

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]interface{}{"success": true, "updated": updatedCount})
}

// Estructura adaptada para recibir el motivo de eliminación
type TreeDeleteRequest struct {
	Level     string `json:"level"` // "comunidad", "torre", "casa"
	Comunidad string `json:"comunidad"`
	Torre     string `json:"torre"`
	Casa      string `json:"casa"`
	Motivo    string `json:"motivo"` // "mudanza", "fallecimiento", "error"
}

const HOJA_EGRESOS = "EGRESOS"

// treeDeleteHandler elimina todos los registros que pertenezcan a la sección seleccionada
func treeDeleteHandler(w http.ResponseWriter, r *http.Request) {
	descargarDeDropbox()
	excelMutex.Lock()
	defer excelMutex.Unlock()

	var req TreeDeleteRequest
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, "JSON inválido", http.StatusBadRequest)
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "No se pudo abrir el archivo Excel", http.StatusInternalServerError)
		return
	}

	rows, _ := f.GetRows(PRIMERA_HOJA)
	if len(rows) < 2 {
		http.Error(w, "El archivo Excel está vacío", http.StatusBadRequest)
		return
	}

	headers := rows[0]
	comIdx, torreIdx, casaIdx, _, _ := getColumnIndices(headers)
	if comIdx == -1 || torreIdx == -1 || casaIdx == -1 {
		http.Error(w, "Columnas de jerarquía requeridas no encontradas", http.StatusInternalServerError)
		return
	}

	var rowsToDelete []int
	var rowsToMigrate [][]string

	for i := 1; i < len(rows); i++ {
		row := rows[i]
		if len(row) <= comIdx || len(row) <= torreIdx || len(row) <= casaIdx {
			continue
		}

		dbCom := strings.TrimSpace(row[comIdx])
		dbTorre := strings.TrimSpace(row[torreIdx])
		dbCasa := strings.TrimSpace(row[casaIdx])

		match := false
		if req.Level == "comunidad" && dbCom == strings.TrimSpace(req.Comunidad) {
			match = true
		} else if req.Level == "torre" && dbCom == strings.TrimSpace(req.Comunidad) && dbTorre == strings.TrimSpace(req.Torre) {
			match = true
		} else if req.Level == "casa" && dbCom == strings.TrimSpace(req.Comunidad) && dbTorre == strings.TrimSpace(req.Torre) && dbCasa == strings.TrimSpace(req.Casa) {
			match = true
		}

		if match {
			rowsToDelete = append(rowsToDelete, i+1)
			rowsToMigrate = append(rowsToMigrate, row)
		}
	}

	// Si el motivo es mudanza o fallecimiento, guardamos en la hoja de EGRESOS
	if (req.Motivo == "mudanza" || req.Motivo == "fallecimiento") && len(rowsToMigrate) > 0 {
		sheetIndex := f.GetSheetIndex(HOJA_EGRESOS)
		if sheetIndex == -1 {
			// Crear la hoja de egresos si no existe
			f.NewSheet(HOJA_EGRESOS)

			// Escribir los encabezados originales + Motivo y Fecha
			egresoHeaders := append(headers, "MOTIVO_EGRESO", "FECHA_EGRESO")
			for colIdx, h := range egresoHeaders {
				cell, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
				f.SetCellValue(HOJA_EGRESOS, cell, h)
			}
		}

		rowsEgresos, _ := f.GetRows(HOJA_EGRESOS)
		nextEgresoRow := len(rowsEgresos) + 1
		fechaHoy := time.Now().Format("02/01/2006 03:04 PM")

		for _, rowData := range rowsToMigrate {
			// Copiar todas las columnas existentes
			for colIdx, val := range rowData {
				cell, _ := excelize.CoordinatesToCellName(colIdx+1, nextEgresoRow)
				f.SetCellValue(HOJA_EGRESOS, cell, val)
			}
			// Añadir motivo y fecha
			cellMotivo, _ := excelize.CoordinatesToCellName(len(headers)+1, nextEgresoRow)
			cellFecha, _ := excelize.CoordinatesToCellName(len(headers)+2, nextEgresoRow)
			f.SetCellValue(HOJA_EGRESOS, cellMotivo, strings.ToUpper(req.Motivo))
			f.SetCellValue(HOJA_EGRESOS, cellFecha, fechaHoy)

			nextEgresoRow++
		}
	}

	// Ordenamos las filas a eliminar en sentido descendente
	for i := 0; i < len(rowsToDelete); i++ {
		for j := i + 1; j < len(rowsToDelete); j++ {
			if rowsToDelete[i] < rowsToDelete[j] {
				rowsToDelete[i], rowsToDelete[j] = rowsToDelete[j], rowsToDelete[i]
			}
		}
	}

	// Remoción de la hoja principal CENSO
	for _, rNum := range rowsToDelete {
		f.RemoveRow(PRIMERA_HOJA, rNum)
	}

	if err := f.Save(); err != nil {
		http.Error(w, "Error al guardar los cambios en el disco", http.StatusInternalServerError)
		return
	}

	go subirADropbox()
	addLog(fmt.Sprintf("Jerarquía: Se procesó la baja por '%s' del nivel %s (Comunidad: %s, Torre: %s, Casa: %s)", req.Motivo, req.Level, req.Comunidad, req.Torre, req.Casa))

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]interface{}{"success": true, "deleted_rows": len(rowsToDelete)})
}

// treeCreateHandler añade una fila con un registro marcador de posición ("Hogar Vacío") para estructurar nuevos niveles
func treeCreateHandler(w http.ResponseWriter, r *http.Request) {
	descargarDeDropbox()
	excelMutex.Lock()
	defer excelMutex.Unlock()

	var req TreeEditRequest
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, "JSON inválido", http.StatusBadRequest)
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "No se pudo abrir el archivo Excel", http.StatusInternalServerError)
		return
	}

	rows, _ := f.GetRows(PRIMERA_HOJA)
	headers := rows[0]
	comIdx, torreIdx, casaIdx, _, nombreIdx := getColumnIndices(headers)
	if comIdx == -1 || torreIdx == -1 || casaIdx == -1 {
		http.Error(w, "Columnas de jerarquía requeridas no encontradas", http.StatusInternalServerError)
		return
	}

	nextRow := len(rows) + 1

	comVal := req.Comunidad
	torreVal := req.Torre
	casaVal := req.Casa

	if req.Level == "comunidad" {
		comVal = req.NewName
		torreVal = "1"
		casaVal = "1"
	} else if req.Level == "torre" {
		torreVal = req.NewName
		casaVal = "1"
	} else if req.Level == "casa" {
		casaVal = req.NewName
	}

	for colIdx := range headers {
		var val string
		switch colIdx {
		case comIdx:
			val = comVal
		case torreIdx:
			val = torreVal
		case casaIdx:
			val = casaVal
		default:
			if colIdx == nombreIdx {
				val = "HOGAR VACIO (CREADO)"
			} else {
				val = ""
			}
		}
		cell, _ := excelize.CoordinatesToCellName(colIdx+1, nextRow)
		f.SetCellValue(PRIMERA_HOJA, cell, val)
	}

	if err := f.Save(); err != nil {
		http.Error(w, "Error al escribir los cambios en el archivo", http.StatusInternalServerError)
		return
	}

	go subirADropbox()
	addLog(fmt.Sprintf("Jerarquía: Se creó el nivel %s (%s - %s - %s)", req.Level, comVal, torreVal, casaVal))

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]interface{}{"success": true})
}

// searchPeopleHandler busca habitantes de forma dinámica comparando nombre o número de cédula
func searchPeopleHandler(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	query := strings.ToLower(strings.TrimSpace(r.URL.Query().Get("q")))
	if query == "" {
		json.NewEncoder(w).Encode([]interface{}{})
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "No se pudo abrir el archivo Excel", http.StatusInternalServerError)
		return
	}

	rows, _ := f.GetRows(PRIMERA_HOJA)
	if len(rows) < 2 {
		json.NewEncoder(w).Encode([]interface{}{})
		return
	}

	comIdx, torreIdx, casaIdx, cedulaIdx, nombreIdx := getColumnIndices(rows[0])
	if nombreIdx == -1 {
		http.Error(w, "Columna de nombre no encontrada", http.StatusInternalServerError)
		return
	}

	type SearchResult struct {
		Nombres   string `json:"nombres"`
		Documento string `json:"documento"`
		Comunidad string `json:"comunidad"`
		Torre     string `json:"torre"`
		Casa      string `json:"casa"`
	}

	var results []SearchResult
	for i := 1; i < len(rows); i++ {
		row := rows[i]
		if len(row) <= nombreIdx {
			continue
		}

		nombre := row[nombreIdx]
		cedula := ""
		if cedulaIdx != -1 && len(row) > cedulaIdx {
			cedula = row[cedulaIdx]
		}

		match := strings.Contains(strings.ToLower(nombre), query) ||
			(cedula != "" && strings.Contains(strings.ToLower(cleanCedula(cedula)), cleanCedula(query)))

		if match {
			com := ""
			if comIdx != -1 && len(row) > comIdx {
				com = row[comIdx]
			}
			tor := ""
			if torreIdx != -1 && len(row) > torreIdx {
				tor = row[torreIdx]
			}
			cas := ""
			if casaIdx != -1 && len(row) > casaIdx {
				cas = row[casaIdx]
			}

			results = append(results, SearchResult{
				Nombres:   nombre,
				Documento: cedula,
				Comunidad: com,
				Torre:     tor,
				Casa:      cas,
			})

			if len(results) >= 10 { // Límite de sugerencias para agilizar la vista
				break
			}
		}
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(results)
}

// Estructuras de datos para eliminaciones en lote
type DeleteTarget struct {
	Level     string `json:"level"` // "comunidad", "torre", "casa"
	Comunidad string `json:"comunidad"`
	Torre     string `json:"torre"`
	Casa      string `json:"casa"`
}

type BulkDeleteRequest struct {
	Targets []DeleteTarget `json:"targets"`
}

// treeDeleteMultipleHandler procesa la eliminación masiva de múltiples targets en una sola transacción
func treeDeleteMultipleHandler(w http.ResponseWriter, r *http.Request) {
	descargarDeDropbox()
	excelMutex.Lock()
	defer excelMutex.Unlock()

	var req BulkDeleteRequest
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, "JSON de lote inválido", http.StatusBadRequest)
		return
	}

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "No se pudo abrir el archivo Excel", http.StatusInternalServerError)
		return
	}

	rows, _ := f.GetRows(PRIMERA_HOJA)
	if len(rows) < 2 {
		http.Error(w, "La base de datos se encuentra vacía", http.StatusBadRequest)
		return
	}

	comIdx, torreIdx, casaIdx, _, _ := getColumnIndices(rows[0])
	if comIdx == -1 || torreIdx == -1 || casaIdx == -1 {
		http.Error(w, "Columnas de jerarquía requeridas no encontradas", http.StatusInternalServerError)
		return
	}

	// Mapa para consolidar y marcar los números de fila únicos a eliminar
	rowsToDeleteMap := make(map[int]bool)

	for i := 1; i < len(rows); i++ {
		row := rows[i]
		if len(row) <= comIdx || len(row) <= torreIdx || len(row) <= casaIdx {
			continue
		}

		dbCom := strings.TrimSpace(row[comIdx])
		dbTorre := strings.TrimSpace(row[torreIdx])
		dbCasa := strings.TrimSpace(row[casaIdx])

		for _, target := range req.Targets {
			match := false
			tCom := strings.TrimSpace(target.Comunidad)
			tTorre := strings.TrimSpace(target.Torre)
			tCasa := strings.TrimSpace(target.Casa)

			if target.Level == "comunidad" && dbCom == tCom {
				match = true
			} else if target.Level == "torre" && dbCom == tCom && dbTorre == tTorre {
				match = true
			} else if target.Level == "casa" && dbCom == tCom && dbTorre == tTorre && dbCasa == tCasa {
				match = true
			}

			if match {
				rowsToDeleteMap[i+1] = true
				break
			}
		}
	}

	// Conversión del mapa a slice para ordenar de forma descendente
	var rowsToDelete []int
	for rNum := range rowsToDeleteMap {
		rowsToDelete = append(rowsToDelete, rNum)
	}

	for i := 0; i < len(rowsToDelete); i++ {
		for j := i + 1; j < len(rowsToDelete); j++ {
			if rowsToDelete[i] < rowsToDelete[j] {
				rowsToDelete[i], rowsToDelete[j] = rowsToDelete[j], rowsToDelete[i]
			}
		}
	}

	// Remoción secuencial segura sin alterar índices superiores
	for _, rNum := range rowsToDelete {
		f.RemoveRow(PRIMERA_HOJA, rNum)
	}

	if err := f.Save(); err != nil {
		http.Error(w, "No se guardaron los cambios de borrado múltiple", http.StatusInternalServerError)
		return
	}

	go subirADropbox()
	addLog(fmt.Sprintf("Jerarquía: Borrado masivo exitoso. Se removieron %d registros", len(rowsToDelete)))

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]interface{}{"success": true, "deleted_rows": len(rowsToDelete)})
}

// Estructura de habitante egresado ajustada a tus cabeceras
type PersonaEgreso struct {
	NombreCompleto string `json:"nombre_completo"`
	Documento      string `json:"documento"`
	Parentesco     string `json:"parentesco"`
	Motivo         string `json:"motivo"`
	Fecha          string `json:"fecha"`
}

type EgresoTreeNode struct {
	Text     string            `json:"text"`
	Type     string            `json:"type"`
	Children []*EgresoTreeNode `json:"children,omitempty"`
	People   []PersonaEgreso   `json:"people,omitempty"`
}

func getEgresosTreeData(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	w.Header().Set("Content-Type", "application/json")

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, `{"error": "No se pudo abrir el Excel: `+err.Error()+`"}`, http.StatusInternalServerError)
		return
	}

	if f.GetSheetIndex(HOJA_EGRESOS) == -1 {
		json.NewEncoder(w).Encode([]*EgresoTreeNode{})
		return
	}

	rows, err := f.GetRows(HOJA_EGRESOS)
	if err != nil || len(rows) < 2 {
		json.NewEncoder(w).Encode([]*EgresoTreeNode{})
		return
	}

	headers := rows[0]
	comIdx, torreIdx, casaIdx, docIdx, nomIdx := getColumnIndices(headers)

	// Detectar dinámicamente parentesco, motivo y fecha
	parIdx, motivoIdx, fechaIdx := -1, -1, -1
	for i, h := range headers {
		clean := strings.TrimSpace(strings.ToLower(h))
		switch clean {
		case "parentesco":
			parIdx = i
		case "motivo_egreso", "motivo":
			motivoIdx = i
		case "fecha_egreso", "fecha":
			fechaIdx = i
		}
	}

	if comIdx == -1 || torreIdx == -1 || casaIdx == -1 {
		http.Error(w, `{"error": "Columnas de jerarquía requeridas no encontradas en EGRESOS"}`, http.StatusBadRequest)
		return
	}

	getVal := func(row []string, idx int) string {
		if idx != -1 && len(row) > idx {
			return strings.TrimSpace(row[idx])
		}
		return ""
	}

	treeMap := make(map[string]map[string]map[string][]PersonaEgreso)

	for _, row := range rows[1:] {
		if len(row) <= comIdx || len(row) <= torreIdx || len(row) <= casaIdx {
			continue
		}

		com := strings.TrimSpace(row[comIdx])
		torre := strings.TrimSpace(row[torreIdx])
		casa := strings.TrimSpace(row[casaIdx])

		if com == "" || torre == "" || casa == "" {
			continue
		}

		p := PersonaEgreso{
			NombreCompleto: getVal(row, nomIdx),
			Documento:      getVal(row, docIdx),
			Parentesco:     getVal(row, parIdx),
			Motivo:         getVal(row, motivoIdx),
			Fecha:          getVal(row, fechaIdx),
		}

		if _, ok := treeMap[com]; !ok {
			treeMap[com] = make(map[string]map[string][]PersonaEgreso)
		}
		if _, ok := treeMap[com][torre]; !ok {
			treeMap[com][torre] = make(map[string][]PersonaEgreso)
		}
		treeMap[com][torre][casa] = append(treeMap[com][torre][casa], p)
	}

	var result []*EgresoTreeNode
	for comName, torres := range treeMap {
		comNode := &EgresoTreeNode{
			Text: comName,
			Type: "comunidad",
		}

		for torreName, casas := range torres {
			torreNode := &EgresoTreeNode{
				Text: "Torre " + torreName,
				Type: "torre",
			}

			for casaName, peopleList := range casas {
				casaNode := &EgresoTreeNode{
					Text:   "Casa/Apto " + casaName,
					Type:   "casa",
					People: peopleList,
				}
				torreNode.Children = append(torreNode.Children, casaNode)
			}
			comNode.Children = append(comNode.Children, torreNode)
		}
		result = append(result, comNode)
	}

	json.NewEncoder(w).Encode(result)
}

// Estructura para la respuesta JSON de estadísticas
type DashboardStats struct {
	TotalPersonas int `json:"total_personas"`
	Hombres       int `json:"hombres"`
	Mujeres       int `json:"mujeres"`
	RTMenores     int `json:"rt_menores"`         // 0 a 17 años
	RTJovenes     int `json:"rt_adultos_jovenes"` // 18 a 35 años
	RTAdultos     int `json:"rt_adultos"`         // 36 a 59 años
	RTAncianos    int `json:"rt_ancianos"`        // 60+ años

	MenoresHombres  int `json:"menores_hombres"`
	MenoresMujeres  int `json:"menores_mujeres"`
	JovenesHombres  int `json:"jovenes_hombres"`
	JovenesMujeres  int `json:"jovenes_mujeres"`
	AdultosHombres  int `json:"adultos_hombres"`
	AdultosMujeres  int `json:"adultos_mujeres"`
	AncianosHombres int `json:"ancianos_hombres"`
	AncianosMujeres int `json:"ancianos_mujeres"`

	Comunidades map[string]int            `json:"comunidades"`
	Torres      map[string]map[string]int `json:"torres"`

	NivelEducativo         map[string]int `json:"nivel_educativo"`
	SituacionLaboral       map[string]int `json:"situacion_laboral"`
	CondicionVivienda      map[string]int `json:"condicion_vivienda"`
	DistribucionParentesco map[string]int `json:"distribucion_parentesco"`
	Enfermedades           map[string]int `json:"enfermedades"`
	Discapacidades         map[string]int `json:"discapacidades"`
	Medicamentos           map[string]int `json:"medicamentos"`

	TotalEgresos  int            `json:"total_egresos"`
	MotivosEgreso map[string]int `json:"motivos_egreso"`
}

func getDashboardStatsHandler(w http.ResponseWriter, r *http.Request) {
	excelMutex.Lock()
	defer excelMutex.Unlock()

	f, err := excelize.OpenFile(EXCEL_FILE)
	if err != nil {
		http.Error(w, "No se pudo abrir la base de datos", http.StatusInternalServerError)
		return
	}

	rows, err := f.GetRows(PRIMERA_HOJA)
	if err != nil || len(rows) < 2 {
		http.Error(w, "La base de datos está vacía", http.StatusInternalServerError)
		return
	}

	headers := rows[0]
	edadIdx, generoIdx, comunidadIdx, torreIdx, nivelEducativoIdx, situacionLaboralIdx, viviendaIdx, parentescoIdx := -1, -1, -1, -1, -1, -1, -1, -1
	enfermedadIdx, discapacidadIdx, medicamentoIdx := -1, -1, -1

	// Buscar las columnas clave
	for i, h := range headers {
		clean := strings.TrimSpace(strings.ToLower(h))
		if clean == "edad" {
			edadIdx = i
		}
		if clean == "genero biológico" || clean == "genero" {
			generoIdx = i
		}
		if clean == "comunidad" {
			comunidadIdx = i
		}
		if clean == "torre" {
			torreIdx = i
		}
		if clean == "nivel educativo" {
			nivelEducativoIdx = i
		}
		if clean == "profesión" {
			situacionLaboralIdx = i
		}
		if clean == "condición de vivienda" || clean == "condicion de vivienda" {
			viviendaIdx = i
		}
		if clean == "parentesco" {
			parentescoIdx = i
		}
		if clean == "enfermedades crónicas" || clean == "enfermedades cronicas" {
			enfermedadIdx = i
		}
		if clean == "tipo de discapacidad" || clean == "tipo de discapacidad" {
			discapacidadIdx = i
		}
		if clean == "medicamento consumido" || clean == "medicamento consumido" {
			medicamentoIdx = i
		}
	}

	var stats DashboardStats
	stats.Comunidades = make(map[string]int)
	stats.Torres = make(map[string]map[string]int)
	stats.NivelEducativo = make(map[string]int)
	stats.SituacionLaboral = make(map[string]int)
	stats.CondicionVivienda = make(map[string]int)
	stats.DistribucionParentesco = make(map[string]int)
	stats.Enfermedades = make(map[string]int)
	stats.Discapacidades = make(map[string]int)
	stats.Medicamentos = make(map[string]int)

	stats.MotivosEgreso = make(map[string]int)

	for i := 1; i < len(rows); i++ {
		row := rows[i]
		stats.TotalPersonas++

		// Conteo por Género
		if generoIdx != -1 && generoIdx < len(row) {
			gen := strings.TrimSpace(strings.ToLower(row[generoIdx]))
			if gen == "masculino" || gen == "m" {
				stats.Hombres++
			} else if gen == "femenino" || gen == "f" {
				stats.Mujeres++
			}
		}

		// Conteo por Rangos de Edad
		// Dentro del bucle que lee las filas:
		if edadIdx != -1 && edadIdx < len(row) {
			edad, err := strconv.Atoi(strings.TrimSpace(row[edadIdx]))
			if err == nil {
				// Detectar género
				isHombre := false
				isMujer := false
				if generoIdx != -1 && generoIdx < len(row) {
					gen := strings.TrimSpace(strings.ToLower(row[generoIdx]))
					if strings.HasPrefix(gen, "masculin") || gen == "m" || gen == "hombre" {
						isHombre = true
					} else if strings.HasPrefix(gen, "femenin") || gen == "f" || gen == "mujer" {
						isMujer = true
					}
				}

				// Clasificación por edad y género
				switch {
				case edad >= 0 && edad <= 17:
					stats.RTMenores++
					if isHombre {
						stats.MenoresHombres++
					}
					if isMujer {
						stats.MenoresMujeres++
					}

				case edad >= 18 && edad <= 35:
					stats.RTJovenes++
					if isHombre {
						stats.JovenesHombres++
					}
					if isMujer {
						stats.JovenesMujeres++
					}

				case edad >= 36 && edad <= 59:
					stats.RTAdultos++
					if isHombre {
						stats.AdultosHombres++
					}
					if isMujer {
						stats.AdultosMujeres++
					}

				case edad >= 60:
					stats.RTAncianos++
					if isHombre {
						stats.AncianosHombres++
					}
					if isMujer {
						stats.AncianosMujeres++
					}
				}
			}
		}

		// Dentro del bucle for i := 1; i < len(rows); i++:
		if comunidadIdx != -1 && comunidadIdx < len(row) {
			comunidad := strings.TrimSpace(row[comunidadIdx])
			torre := "1"
			if torreIdx != -1 && torreIdx < len(row) && strings.TrimSpace(row[torreIdx]) != "" {
				torre = strings.TrimSpace(row[torreIdx])
			}

			if comunidad != "" {
				stats.Comunidades[comunidad]++

				if _, ok := stats.Torres[comunidad]; !ok {
					stats.Torres[comunidad] = make(map[string]int)
				}
				stats.Torres[comunidad][torre]++
			}
		}

		// Conteo por Nivel Educativo
		if nivelEducativoIdx != -1 && nivelEducativoIdx < len(row) {
			nivel := strings.TrimSpace(row[nivelEducativoIdx])
			if nivel != "" {
				stats.NivelEducativo[nivel]++
			}
		}

		// Conteo por Situación Laboral
		if situacionLaboralIdx != -1 && situacionLaboralIdx < len(row) {
			situacion := strings.TrimSpace(row[situacionLaboralIdx])
			if situacion != "" {
				stats.SituacionLaboral[situacion]++
			}
		}

		// Conteo por Condición de Vivienda
		if viviendaIdx != -1 && viviendaIdx < len(row) {
			vivienda := strings.TrimSpace(row[viviendaIdx])
			if vivienda != "" {
				stats.CondicionVivienda[vivienda]++
			}
		}

		// Conteo por Distribución de Parentesco
		if parentescoIdx != -1 && parentescoIdx < len(row) {
			parentesco := strings.TrimSpace(row[parentescoIdx])
			if parentesco != "" {
				stats.DistribucionParentesco[parentesco]++
			}
		}

		// Conteo por Enfermedades
		if enfermedadIdx != -1 && enfermedadIdx < len(row) {
			enfermedad := strings.TrimSpace(row[enfermedadIdx])
			if enfermedad != "" {
				stats.Enfermedades[enfermedad]++
			}
		}

		// Conteo por Discapacidades
		if discapacidadIdx != -1 && discapacidadIdx < len(row) {
			discapacidad := strings.TrimSpace(row[discapacidadIdx])
			if discapacidad != "" {
				stats.Discapacidades[discapacidad]++
			}
		}

		// Conteo por Medicamentos
		if medicamentoIdx != -1 && medicamentoIdx < len(row) {
			medicamento := strings.TrimSpace(row[medicamentoIdx])
			if medicamento != "" {
				stats.Medicamentos[medicamento]++
			}
		}

	}

	// =========================================================
	// LECTURA DE LA SEGUNDA HOJA: "EGRESOS"
	// =========================================================
	egresosRows, err := f.GetRows("EGRESOS")
	if err == nil && len(egresosRows) > 1 {
		headersEgresos := egresosRows[0]
		motivoIdx := -1

		for i, h := range headersEgresos {
			clean := strings.TrimSpace(strings.ToUpper(h))
			if clean == "MOTIVO_EGRESO" {
				motivoIdx = i
				break
			}
		}

		if motivoIdx != -1 {
			for i := 1; i < len(egresosRows); i++ {
				row := egresosRows[i]
				val := ""
				if motivoIdx < len(row) {
					val = strings.TrimSpace(strings.ToUpper(row[motivoIdx]))
				}
				if val != "" {
					stats.MotivosEgreso[val]++
					stats.TotalEgresos++
				}
			}
		}
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(stats)
}
