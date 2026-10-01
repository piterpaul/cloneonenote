package com.cloneonenote.feature.editor

import androidx.compose.animation.AnimatedVisibility
import androidx.compose.animation.expandVertically
import androidx.compose.animation.fadeIn
import androidx.compose.animation.fadeOut
import androidx.compose.animation.shrinkVertically
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.horizontalScroll
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.lazy.rememberLazyListState
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.CircleShape
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.text.BasicTextField
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.SolidColor
import androidx.compose.ui.text.AnnotatedString
import androidx.compose.ui.text.SpanStyle
import androidx.compose.ui.text.TextStyle
import androidx.compose.ui.text.buildAnnotatedString
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.OffsetMapping
import androidx.compose.ui.text.input.TransformedText
import androidx.compose.ui.text.input.VisualTransformation
import androidx.compose.ui.text.style.TextDecoration
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import kotlinx.coroutines.delay
import kotlinx.coroutines.flow.MutableSharedFlow
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.SharedFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asSharedFlow
import kotlinx.coroutines.flow.asStateFlow
import kotlinx.coroutines.flow.update
import kotlinx.coroutines.launch
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale
import java.util.UUID

// ============================================================================
// 1. MODELOS DE DOMINIO Y ESTADO DE UI (CLEAN ARCHITECTURE)
// ============================================================================

data class NoteDocument(
    val id: String = UUID.randomUUID().toString(),
    val notebookName: String = "Mi Bloc Principal",
    val sectionName: String = "Notas de Proyecto",
    val title: String = "Reunión de Arquitectura y Hoja de Ruta Q4",
    val content: String = """
        En la reunión de hoy decidimos migrar el módulo de sincronización a Google Drive API v3 antes del viernes.
        
        Acuerdos principales del equipo:
        - Revisar el diseño responsivo en tablets Android y Chromebooks.
        - Preparar la demo del Boli Táctil con sensibilidad a la presión.
        - Pendiente: Enviar el informe de rendimiento al equipo de producto.
        - Recordar: Configurar las alertas de copia de seguridad automática mañana a las 09:00.
    """.trimIndent(),
    val updatedAt: String = "Hoy, 10:30"
)

data class TextMatchRange(
    val start: Int,
    val end: Int,
    val lineIndex: Int
)

enum class ChatRole { USER, ASSISTANT }

data class AiChatMessage(
    val id: String = UUID.randomUUID().toString(),
    val role: ChatRole,
    val text: String,
    val suggestedTasks: List<ExtractedNoteTask> = emptyList(),
    val timestamp: String = SimpleDateFormat("HH:mm", Locale.getDefault()).format(Date())
)

data class ExtractedNoteTask(
    val id: String = UUID.randomUUID().toString(),
    val description: String,
    val sourceNoteTitle: String,
    val isCompleted: Boolean = false,
    val reminderTimeLabel: String? = null
)

data class NoteEditorUiState(
    val activeNote: NoteDocument = NoteDocument(),
    // Estado 1: Búsqueda Interna ("Buscar en esta nota/bloc")
    val isSearchVisible: Boolean = false,
    val searchQuery: String = "",
    val searchMatches: List<TextMatchRange> = emptyList(),
    val currentMatchIndex: Int = -1,
    // Estado 2: Agente de IA Conversacional ("Asistente de Nota")
    val isAiSheetOpen: Boolean = false,
    val isAiThinking: Boolean = false,
    val aiMessages: List<AiChatMessage> = emptyList(),
    val extractedTasks: List<ExtractedNoteTask> = emptyList(),
    val localReminders: List<ExtractedNoteTask> = emptyList()
)

// ============================================================================
// 2. VIEWMODEL: BÚSQUEDA INTERNA EN TIEMPO REAL + AGENTE IA CONVERSACIONAL
// ============================================================================

class NoteEditorViewModel : ViewModel() {

    private val _uiState = MutableStateFlow(NoteEditorUiState())
    val uiState: StateFlow<NoteEditorUiState> = _uiState.asStateFlow()

    private val _snackbarEvents = MutableSharedFlow<String>()
    val snackbarEvents: SharedFlow<String> = _snackbarEvents.asSharedFlow()

    init {
        seedWelcomeAiMessage()
        extractTasksFromNote(silent = true)
    }

    // --- Edición de la Nota ---

    fun onNoteTitleChanged(newTitle: String) {
        _uiState.update { it.copy(activeNote = it.activeNote.copy(title = newTitle)) }
    }

    fun onNoteContentChanged(newContent: String) {
        _uiState.update { it.copy(activeNote = it.activeNote.copy(content = newContent)) }
        if (_uiState.value.isSearchVisible && _uiState.value.searchQuery.isNotBlank()) {
            recalculateSearchMatches(_uiState.value.searchQuery, resetIndex = false)
        }
    }

    // --- Funcionalidad 1: Búsqueda Interna ("Buscar en esta nota/bloc") ---

    fun toggleSearchBar() {
        val nextVisible = !_uiState.value.isSearchVisible
        if (!nextVisible) {
            _uiState.update {
                it.copy(
                    isSearchVisible = false,
                    searchQuery = "",
                    searchMatches = emptyList(),
                    currentMatchIndex = -1
                )
            }
        } else {
            _uiState.update { it.copy(isSearchVisible = true) }
        }
    }

    fun onSearchQueryChanged(query: String) {
        recalculateSearchMatches(query, resetIndex = true)
    }

    private fun recalculateSearchMatches(query: String, resetIndex: Boolean) {
        val trimmed = query.trim()
        if (trimmed.isEmpty()) {
            _uiState.update {
                it.copy(
                    searchQuery = query,
                    searchMatches = emptyList(),
                    currentMatchIndex = -1
                )
            }
            return
        }

        val content = _uiState.value.activeNote.content
        val matches = mutableListOf<TextMatchRange>()
        var searchStart = 0

        while (searchStart < content.length) {
            val foundIndex = content.indexOf(trimmed, startIndex = searchStart, ignoreCase = true)
            if (foundIndex == -1) break
            val lineIdx = content.substring(0, foundIndex).count { it == '\n' }
            matches.add(
                TextMatchRange(
                    start = foundIndex,
                    end = foundIndex + trimmed.length,
                    lineIndex = lineIdx
                )
            )
            searchStart = foundIndex + trimmed.length
        }

        val nextIndex = when {
            matches.isEmpty() -> -1
            resetIndex || _uiState.value.currentMatchIndex !in matches.indices -> 0
            else -> _uiState.value.currentMatchIndex
        }

        _uiState.update {
            it.copy(
                searchQuery = query,
                searchMatches = matches,
                currentMatchIndex = nextIndex
            )
        }
    }

    fun navigateToNextMatch() {
        val matches = _uiState.value.searchMatches
        if (matches.isEmpty()) return
        val next = (_uiState.value.currentMatchIndex + 1) % matches.size
        _uiState.update { it.copy(currentMatchIndex = next) }
    }

    fun navigateToPreviousMatch() {
        val matches = _uiState.value.searchMatches
        if (matches.isEmpty()) return
        val prev = (_uiState.value.currentMatchIndex - 1 + matches.size) % matches.size
        _uiState.update { it.copy(currentMatchIndex = prev) }
    }

    // --- Funcionalidad 2: Agente de IA Conversacional ("Asistente de Nota") ---

    fun setAiSheetOpen(open: Boolean) {
        _uiState.update { it.copy(isAiSheetOpen = open) }
        if (open) {
            extractTasksFromNote(silent = true)
        }
    }

    private fun seedWelcomeAiMessage() {
        val note = _uiState.value.activeNote
        val welcome = AiChatMessage(
            role = ChatRole.ASSISTANT,
            text = "👋 ¡Hola! Soy tu Asistente de Nota IA. Tengo como contexto automático tu nota actual «${note.title}» del bloc «${note.notebookName}». Puedes preguntarme qué se decidió en tus apuntes o pedirme extraer tareas y recordatorios rápidos."
        )
        _uiState.update { it.copy(aiMessages = listOf(welcome)) }
    }

    fun sendMessageToAi(userPrompt: String) {
        val cleanPrompt = userPrompt.trim()
        if (cleanPrompt.isEmpty()) return

        val userMsg = AiChatMessage(role = ChatRole.USER, text = cleanPrompt)
        _uiState.update {
            it.copy(
                aiMessages = it.aiMessages + userMsg,
                isAiThinking = true
            )
        }

        viewModelScope.launch {
            delay(450) // Simulación de latencia de inferencia del Agente IA
            val note = _uiState.value.activeNote
            val tasks = extractTasksFromNote(silent = true)
            val response = buildContextualAiResponse(cleanPrompt, note, tasks)

            _uiState.update {
                it.copy(
                    aiMessages = it.aiMessages + response,
                    isAiThinking = false
                )
            }
        }
    }

    fun extractTasksFromNote(silent: Boolean = false): List<ExtractedNoteTask> {
        val note = _uiState.value.activeNote
        val actionKeywords = listOf(
            "pendiente", "revisar", "preparar", "enviar", "recordar",
            "configurar", "migrar", "hacer", "llamar", "entregar"
        )

        val tasks = note.content
            .lines()
            .map { it.trim().removePrefix("-").removePrefix("•").trim() }
            .filter { line ->
                line.length > 8 && actionKeywords.any { kw -> line.contains(kw, ignoreCase = true) }
            }
            .map { line ->
                ExtractedNoteTask(
                    description = line,
                    sourceNoteTitle = note.title
                )
            }

        _uiState.update { it.copy(extractedTasks = tasks) }
        if (!silent) {
            viewModelScope.launch {
                _snackbarEvents.emit("✨ Se identificaron ${tasks.size} tareas en la nota actual.")
            }
        }
        return tasks
    }

    fun scheduleQuickReminder(task: ExtractedNoteTask, presetLabel: String = "Mañana 09:00") {
        val scheduledTask = task.copy(reminderTimeLabel = presetLabel)
        _uiState.update { state ->
            val updatedReminders = (listOf(scheduledTask) + state.localReminders)
                .distinctBy { it.description }
            state.copy(localReminders = updatedReminders)
        }

        viewModelScope.launch {
            _snackbarEvents.emit("🔔 Recordatorio guardado ($presetLabel): «${task.description}»")
        }
    }

    private fun buildContextualAiResponse(
        query: String,
        note: NoteDocument,
        tasks: List<ExtractedNoteTask>
    ): AiChatMessage {
        val q = query.lowercase(Locale.getDefault())

        return when {
            q.contains("decid") || q.contains("reunión") || q.contains("acuerdo") -> {
                val relevantLines = note.content.lines().filter { it.isNotBlank() }.take(4)
                AiChatMessage(
                    role = ChatRole.ASSISTANT,
                    text = "📌 Según tu nota «${note.title}», esto fue lo que decidisteis:\n\n" +
                        relevantLines.joinToString("\n") { "• ${it.trim()}" },
                    suggestedTasks = tasks.take(2)
                )
            }
            q.contains("tarea") || q.contains("pendiente") || q.contains("compromiso") || q.contains("recordatorio") -> {
                AiChatMessage(
                    role = ChatRole.ASSISTANT,
                    text = "📋 He analizado el contenido de «${note.title}» y he extraído ${tasks.size} compromisos o tareas pendientes. Puedes configurar un recordatorio rápido con un toque:",
                    suggestedTasks = tasks
                )
            }
            else -> {
                val matchingLines = note.content.lines().filter { line ->
                    query.split(" ").filter { it.length > 3 }.any { kw ->
                        line.contains(kw, ignoreCase = true)
                    }
                }
                val summaryText = if (matchingLines.isNotEmpty()) {
                    "🔎 Encontré estas referencias en tu nota «${note.title}»:\n\n" +
                        matchingLines.joinToString("\n") { "“${it.trim()}”" }
                } else {
                    "📝 Resumen de «${note.title}» (${note.sectionName}):\n" +
                        note.content.take(220) + "..."
                }
                AiChatMessage(
                    role = ChatRole.ASSISTANT,
                    text = summaryText,
                    suggestedTasks = tasks.take(2)
                )
            }
        }
    }
}

// ============================================================================
// 3. COMPONENTES VISUALES EN JETPACK COMPOSE (MATERIAL DESIGN 3)
// ============================================================================

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun OneNoteEditorScreen(
    viewModel: NoteEditorViewModel,
    modifier: Modifier = Modifier
) {
    val uiState by viewModel.uiState.collectAsState()
    val snackbarHostState = remember { SnackbarHostState() }
    val editorScrollState = rememberScrollState()

    // Escuchar eventos de recordatorios y notificaciones locales
    LaunchedEffect(Unit) {
        viewModel.snackbarEvents.collect { message ->
            snackbarHostState.showSnackbar(
                message = message,
                duration = SnackbarDuration.Short
            )
        }
    }

    // Desplazamiento automático hacia la coincidencia activa en la búsqueda interna
    LaunchedEffect(uiState.currentMatchIndex, uiState.searchMatches) {
        val activeMatch = uiState.searchMatches.getOrNull(uiState.currentMatchIndex)
        if (activeMatch != null) {
            val targetPx = (activeMatch.lineIndex * 56).coerceAtLeast(0)
            editorScrollState.animateScrollTo(targetPx)
        }
    }

    Scaffold(
        modifier = modifier.fillMaxSize(),
        snackbarHost = { SnackbarHost(hostState = snackbarHostState) },
        topBar = {
            Column {
                // TopAppBar estilo Microsoft OneNote con Botón de Lupa e IA
                TopAppBar(
                    title = {
                        Column {
                            Text(
                                text = uiState.activeNote.notebookName,
                                style = MaterialTheme.typography.titleMedium,
                                fontWeight = FontWeight.Bold,
                                color = Color.White
                            )
                            Text(
                                text = uiState.activeNote.sectionName,
                                style = MaterialTheme.typography.labelSmall,
                                color = Color.White.copy(alpha = 0.85f)
                            )
                        }
                    },
                    actions = {
                        // 1. Botón de Búsqueda Interna ("Buscar en esta nota/bloc")
                        IconButton(onClick = { viewModel.toggleSearchBar() }) {
                            Icon(
                                imageVector = Icons.Default.Search,
                                contentDescription = "Buscar en esta nota",
                                tint = Color.White
                            )
                        }
                        // 2. Acceso rápido en TopAppBar al Asistente de IA
                        IconButton(onClick = { viewModel.setAiSheetOpen(true) }) {
                            Icon(
                                imageVector = Icons.Default.AutoAwesome,
                                contentDescription = "Abrir Asistente de Nota IA",
                                tint = Color(0xFFFDE047)
                            )
                        }
                    },
                    colors = TopAppBarDefaults.topAppBarColors(
                        containerColor = Color(0xFF7719AA) // Púrpura OneNote
                    )
                )

                // Barra desplegable de Búsqueda Interna con resaltado en tiempo real
                AnimatedVisibility(
                    visible = uiState.isSearchVisible,
                    enter = expandVertically() + fadeIn(),
                    exit = shrinkVertically() + fadeOut()
                ) {
                    InternalNoteSearchBar(
                        query = uiState.searchQuery,
                        matchCount = uiState.searchMatches.size,
                        currentMatchIndex = uiState.currentMatchIndex,
                        onQueryChange = viewModel::onSearchQueryChanged,
                        onPreviousMatch = viewModel::navigateToPreviousMatch,
                        onNextMatch = viewModel::navigateToNextMatch,
                        onClose = viewModel::toggleSearchBar
                    )
                }
            }
        },
        floatingActionButton = {
            // 2. Floating Action Button (FAB) para el Agente de IA Conversacional
            ExtendedFloatingActionButton(
                onClick = { viewModel.setAiSheetOpen(true) },
                containerColor = Color(0xFF7719AA),
                contentColor = Color.White,
                icon = {
                    Icon(
                        imageVector = Icons.Default.AutoAwesome,
                        contentDescription = null,
                        tint = Color(0xFFFDE047)
                    )
                },
                text = {
                    Text(
                        text = "Asistente de Nota",
                        fontWeight = FontWeight.SemiBold
                    )
                }
            )
        }
    ) { innerPadding ->
        // Área de Edición del Bloc de Notas con Resaltado Dinámico de Coincidencias
        Column(
            modifier = Modifier
                .fillMaxSize()
                .padding(innerPadding)
                .verticalScroll(editorScrollState)
                .padding(horizontal = 20.dp, vertical = 16.dp)
        ) {
            // Título de la Nota
            OutlinedTextField(
                value = uiState.activeNote.title,
                onValueChange = viewModel::onNoteTitleChanged,
                textStyle = MaterialTheme.typography.headlineSmall.copy(fontWeight = FontWeight.Bold),
                placeholder = { Text("Título de la página...") },
                modifier = Modifier.fillMaxWidth(),
                colors = OutlinedTextFieldDefaults.colors(
                    focusedBorderColor = Color(0xFF7719AA),
                    unfocusedBorderColor = Color.Transparent
                )
            )

            Text(
                text = "📅 ${uiState.activeNote.updatedAt} · Contexto activo para IA",
                style = MaterialTheme.typography.labelMedium,
                color = MaterialTheme.colorScheme.onSurfaceVariant,
                modifier = Modifier.padding(horizontal = 4.dp, vertical = 4.dp)
            )

            HorizontalDivider(modifier = Modifier.padding(vertical = 8.dp))

            // Editor de Texto con VisualTransformation para resaltar coincidencias en tiempo real
            HighlightedNoteEditorField(
                content = uiState.activeNote.content,
                matches = uiState.searchMatches,
                activeMatchIndex = uiState.currentMatchIndex,
                onContentChange = viewModel::onNoteContentChanged,
                modifier = Modifier
                    .fillMaxWidth()
                    .defaultMinSize(minHeight = 480.dp)
            )
        }

        // BottomSheet Material 3 del Agente de IA Conversacional ("Asistente de Nota")
        if (uiState.isAiSheetOpen) {
            AiAssistantBottomSheet(
                uiState = uiState,
                onDismiss = { viewModel.setAiSheetOpen(false) },
                onSendPrompt = viewModel::sendMessageToAi,
                onExtractTasks = { viewModel.extractTasksFromNote(silent = false) },
                onScheduleReminder = viewModel::scheduleQuickReminder
            )
        }
    }
}

/**
 * Componente 1: Barra de Búsqueda Interna con controles "Anterior" / "Siguiente"
 */
@Composable
fun InternalNoteSearchBar(
    query: String,
    matchCount: Int,
    currentMatchIndex: Int,
    onQueryChange: (String) -> Unit,
    onPreviousMatch: () -> Unit,
    onNextMatch: () -> Unit,
    onClose: () -> Unit,
    modifier: Modifier = Modifier
) {
    Surface(
        modifier = modifier.fillMaxWidth(),
        color = MaterialTheme.colorScheme.surfaceContainerHigh,
        tonalElevation = 4.dp,
        shadowElevation = 4.dp
    ) {
        Row(
            modifier = Modifier
                .fillMaxWidth()
                .padding(horizontal = 12.dp, vertical = 8.dp),
            verticalAlignment = Alignment.CenterVertically,
            horizontalArrangement = Arrangement.spacedBy(8.dp)
        ) {
            OutlinedTextField(
                value = query,
                onValueChange = onQueryChange,
                modifier = Modifier.weight(1f),
                singleLine = true,
                placeholder = { Text("Buscar en esta nota/bloc...") },
                leadingIcon = {
                    Icon(
                        imageVector = Icons.Default.Search,
                        contentDescription = null,
                        tint = Color(0xFF7719AA)
                    )
                },
                shape = RoundedCornerShape(24.dp)
            )

            // Indicador de coincidencias (ej. "1 / 4")
            Surface(
                shape = RoundedCornerShape(16.dp),
                color = Color(0xFFF3E8FF)
            ) {
                Text(
                    text = if (matchCount > 0) "${currentMatchIndex + 1} / $matchCount" else "0 / 0",
                    style = MaterialTheme.typography.labelMedium,
                    fontWeight = FontWeight.Bold,
                    color = Color(0xFF7719AA),
                    modifier = Modifier.padding(horizontal = 10.dp, vertical = 6.dp)
                )
            }

            // Botones "Anterior" y "Siguiente"
            IconButton(
                onClick = onPreviousMatch,
                enabled = matchCount > 0
            ) {
                Icon(
                    imageVector = Icons.Default.KeyboardArrowUp,
                    contentDescription = "Coincidencia anterior"
                )
            }

            IconButton(
                onClick = onNextMatch,
                enabled = matchCount > 0
            ) {
                Icon(
                    imageVector = Icons.Default.KeyboardArrowDown,
                    contentDescription = "Coincidencia siguiente"
                )
            }

            IconButton(onClick = onClose) {
                Icon(
                    imageVector = Icons.Default.Close,
                    contentDescription = "Cerrar búsqueda"
                )
            }
        }
    }
}

/**
 * Editor de texto que aplica resaltado (highlight) en tiempo real sobre las coincidencias
 * sin alterar la posición del cursor mediante VisualTransformation.
 */
@Composable
fun HighlightedNoteEditorField(
    content: String,
    matches: List<TextMatchRange>,
    activeMatchIndex: Int,
    onContentChange: (String) -> Unit,
    modifier: Modifier = Modifier
) {
    val highlightTransformation = remember(matches, activeMatchIndex) {
        VisualTransformation { originalText ->
            val annotated = buildAnnotatedString {
                append(originalText.text)
                matches.forEachIndexed { index, range ->
                    if (range.start >= 0 && range.end <= originalText.length) {
                        val isCurrent = index == activeMatchIndex
                        addStyle(
                            style = SpanStyle(
                                background = if (isCurrent) Color(0xFFF97316) else Color(0xFFFEF08A),
                                color = if (isCurrent) Color.White else Color(0xFF0F172A),
                                fontWeight = if (isCurrent) FontWeight.Bold else FontWeight.Normal,
                                textDecoration = if (isCurrent) TextDecoration.Underline else TextDecoration.None
                            ),
                            start = range.start,
                            end = range.end
                        )
                    }
                }
            }
            TransformedText(annotated, OffsetMapping.Identity)
        }
    }

    BasicTextField(
        value = content,
        onValueChange = onContentChange,
        modifier = modifier.padding(8.dp),
        textStyle = TextStyle(
            fontSize = 16.sp,
            lineHeight = 26.sp,
            color = MaterialTheme.colorScheme.onSurface
        ),
        cursorBrush = SolidColor(Color(0xFF7719AA)),
        visualTransformation = highlightTransformation
    )
}

/**
 * Componente 2: Hoja Inferior (ModalBottomSheet M3) del Agente de IA Conversacional ("Asistente de Nota")
 */
@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun AiAssistantBottomSheet(
    uiState: NoteEditorUiState,
    onDismiss: () -> Unit,
    onSendPrompt: (String) -> Unit,
    onExtractTasks: () -> Unit,
    onScheduleReminder: (ExtractedNoteTask, String) -> Unit
) {
    val sheetState = rememberModalBottomSheetState(skipPartiallyExpanded = true)
    var inputPrompt by remember { mutableStateOf("") }
    var selectedTab by remember { mutableIntStateOf(0) }
    val chatListState = rememberLazyListState()

    LaunchedEffect(uiState.aiMessages.size) {
        if (uiState.aiMessages.isNotEmpty()) {
            chatListState.animateScrollToItem(uiState.aiMessages.lastIndex)
        }
    }

    ModalBottomSheet(
        onDismissRequest = onDismiss,
        sheetState = sheetState,
        shape = RoundedCornerShape(topStart = 24.dp, topEnd = 24.dp)
    ) {
        Column(
            modifier = Modifier
                .fillMaxWidth()
                .fillMaxHeight(0.82f)
                .padding(horizontal = 16.dp)
        ) {
            // Cabecera con indicador de Contexto Activo
            Row(
                modifier = Modifier.fillMaxWidth(),
                verticalAlignment = Alignment.CenterVertically,
                horizontalArrangement = Arrangement.SpaceBetween
            ) {
                Row(
                    verticalAlignment = Alignment.CenterVertically,
                    horizontalArrangement = Arrangement.spacedBy(10.dp)
                ) {
                    Box(
                        modifier = Modifier
                            .size(38.dp)
                            .clip(CircleShape)
                            .background(Color(0xFF7719AA)),
                        contentAlignment = Alignment.Center
                    ) {
                        Icon(
                            imageVector = Icons.Default.AutoAwesome,
                            contentDescription = null,
                            tint = Color(0xFFFDE047)
                        )
                    }
                    Column {
                        Text(
                            text = "Asistente de Nota IA",
                            style = MaterialTheme.typography.titleMedium,
                            fontWeight = FontWeight.Bold
                        )
                        Text(
                            text = "Contexto activo: «${uiState.activeNote.title}»",
                            style = MaterialTheme.typography.labelSmall,
                            color = MaterialTheme.colorScheme.onSurfaceVariant,
                            maxLines = 1,
                            overflow = TextOverflow.Ellipsis
                        )
                    }
                }

                IconButton(onClick = onDismiss) {
                    Icon(Icons.Default.Close, contentDescription = "Cerrar Asistente")
                }
            }

            Spacer(modifier = Modifier.height(8.dp))

            // Selector de Pestañas: Chat Contextual vs Tareas y Recordatorios
            PrimaryTabRow(selectedTabIndex = selectedTab) {
                Tab(
                    selected = selectedTab == 0,
                    onClick = { selectedTab = 0 },
                    text = { Text("💬 Chat con la Nota") }
                )
                Tab(
                    selected = selectedTab == 1,
                    onClick = { selectedTab = 1 },
                    text = { Text("☑️ Tareas (${uiState.extractedTasks.size})") }
                )
            }

            if (selectedTab == 0) {
                // Sugerencias rápidas
                Row(
                    modifier = Modifier
                        .fillMaxWidth()
                        .horizontalScroll(rememberScrollState())
                        .padding(vertical = 8.dp),
                    horizontalArrangement = Arrangement.spacedBy(8.dp)
                ) {
                    AssistChip(
                        onClick = { onSendPrompt("¿Qué decidimos en la reunión según esta nota?") },
                        label = { Text("❓ ¿Qué decidimos en la reunión?") }
                    )
                    AssistChip(
                        onClick = { onSendPrompt("Extrae las tareas pendientes y configura recordatorios") },
                        label = { Text("📋 Extraer tareas pendientes") }
                    )
                }

                // Lista de mensajes del chat
                LazyColumn(
                    state = chatListState,
                    modifier = Modifier
                        .weight(1f)
                        .fillMaxWidth(),
                    verticalArrangement = Arrangement.spacedBy(10.dp)
                ) {
                    items(uiState.aiMessages, key = { it.id }) { msg ->
                        AiChatBubble(
                            message = msg,
                            onScheduleQuickReminder = { task ->
                                onScheduleReminder(task, "Mañana 09:00")
                            }
                        )
                    }
                }

                // Barra de entrada de preguntas en lenguaje natural
                Row(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(vertical = 12.dp),
                    verticalAlignment = Alignment.CenterVertically,
                    horizontalArrangement = Arrangement.spacedBy(8.dp)
                ) {
                    OutlinedTextField(
                        value = inputPrompt,
                        onValueChange = { inputPrompt = it },
                        modifier = Modifier.weight(1f),
                        placeholder = { Text("Pregunta sobre tus apuntes...") },
                        singleLine = true,
                        shape = RoundedCornerShape(24.dp)
                    )
                    FilledIconButton(
                        onClick = {
                            onSendPrompt(inputPrompt)
                            inputPrompt = ""
                        },
                        colors = IconButtonDefaults.filledIconButtonColors(
                            containerColor = Color(0xFF7719AA)
                        )
                    ) {
                        Icon(
                            imageVector = Icons.Default.Send,
                            contentDescription = "Enviar pregunta",
                            tint = Color.White
                        )
                    }
                }
            } else {
                // Pestaña de Gestión de Tareas y Recordatorios Rápidos
                Column(
                    modifier = Modifier
                        .fillMaxSize()
                        .padding(vertical = 12.dp),
                    verticalArrangement = Arrangement.spacedBy(10.dp)
                ) {
                    Button(
                        onClick = onExtractTasks,
                        modifier = Modifier.fillMaxWidth(),
                        colors = ButtonDefaults.buttonColors(containerColor = Color(0xFF7719AA))
                    ) {
                        Icon(Icons.Default.AutoAwesome, contentDescription = null)
                        Spacer(modifier = Modifier.width(8.dp))
                        Text("Analizar Nota y Extraer Tareas Pendientes")
                    }

                    LazyColumn(
                        modifier = Modifier.weight(1f),
                        verticalArrangement = Arrangement.spacedBy(8.dp)
                    ) {
                        items(uiState.extractedTasks, key = { it.id }) { task ->
                            TaskReminderItemCard(
                                task = task,
                                onSetReminder = { preset -> onScheduleReminder(task, preset) }
                            )
                        }
                    }
                }
            }
        }
    }
}

@Composable
private fun AiChatBubble(
    message: AiChatMessage,
    onScheduleQuickReminder: (ExtractedNoteTask) -> Unit
) {
    val isUser = message.role == ChatRole.USER
    Column(
        modifier = Modifier.fillMaxWidth(),
        horizontalAlignment = if (isUser) Alignment.End else Alignment.Start
    ) {
        Surface(
            shape = RoundedCornerShape(16.dp),
            color = if (isUser) Color(0xFF7719AA) else MaterialTheme.colorScheme.surfaceContainerHigh,
            modifier = Modifier.widthIn(max = 320.dp)
        ) {
            Column(modifier = Modifier.padding(12.dp)) {
                Text(
                    text = message.text,
                    color = if (isUser) Color.White else MaterialTheme.colorScheme.onSurface,
                    style = MaterialTheme.typography.bodyMedium
                )

                if (message.suggestedTasks.isNotEmpty()) {
                    Spacer(modifier = Modifier.height(8.dp))
                    message.suggestedTasks.forEach { task ->
                        OutlinedButton(
                            onClick = { onScheduleQuickReminder(task) },
                            modifier = Modifier
                                .fillMaxWidth()
                                .padding(top = 4.dp)
                        ) {
                            Icon(Icons.Default.NotificationsActive, contentDescription = null)
                            Spacer(modifier = Modifier.width(6.dp))
                            Text(
                                text = "Recordar: ${task.description}",
                                maxLines = 1,
                                overflow = TextOverflow.Ellipsis
                            )
                        }
                    }
                }
            }
        }
    }
}

@Composable
private fun TaskReminderItemCard(
    task: ExtractedNoteTask,
    onSetReminder: (String) -> Unit
) {
    ElevatedCard(modifier = Modifier.fillMaxWidth()) {
        Column(modifier = Modifier.padding(12.dp)) {
            Text(
                text = task.description,
                style = MaterialTheme.typography.bodyMedium,
                fontWeight = FontWeight.SemiBold
            )
            Text(
                text = "Extraído de: ${task.sourceNoteTitle}",
                style = MaterialTheme.typography.labelSmall,
                color = MaterialTheme.colorScheme.onSurfaceVariant
            )
            Spacer(modifier = Modifier.height(8.dp))
            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                AssistChip(
                    onClick = { onSetReminder("En 15 min") },
                    label = { Text("⏰ En 15 min") }
                )
                AssistChip(
                    onClick = { onSetReminder("Mañana 09:00") },
                    label = { Text("🔔 Mañana 09:00") }
                )
            }
        }
    }
}
