import React, { useState, useEffect, useMemo } from 'react';
import { UploadCloud, LayoutGrid, Network, Download, Users, Briefcase, ChevronRight, ChevronDown, User, Play, Building, Plus, Minus } from 'lucide-react';

// Mock Data for Demo Purposes
const MOCK_DATA = [
  { Email: "egr@vortx.com.br", Name: "Everson Gonçalves Ramos", "Job Title": "Diretor de Tecnologia", "Manager Name": "", "Manager Email": "", Organization: "Vortx", Department: "DIRETORIA" },
  { Email: "dap@vortx.com.br", Name: "Diego Fellipe Antunes Pedrosa", "Job Title": "Head de Infra. e Segurança da Informação", "Manager Name": "Everson Gonçalves Ramos", "Manager Email": "egr@vortx.com.br", Organization: "Vortx", Department: "PRODUTO & TECNOLOGIA" },
  { Email: "ggm@vortx.com.br", Name: "Giovanni Gomes Miron", "Job Title": "Team Leader de Segurança da Informação", "Manager Name": "Diego Fellipe Antunes Pedrosa", "Manager Email": "dap@vortx.com.br", Organization: "Vortx", Department: "SEGURANÇA DA INFORMAÇÃO" },
  { Email: "mpp@vortx.com.br", Name: "Matheus Picoli", "Job Title": "Analista de Segurança", "Manager Name": "Giovanni Gomes Miron", "Manager Email": "ggm@vortx.com.br", Organization: "Vortx", Department: "SEGURANÇA DA INFORMAÇÃO" },
  { Email: "fgc@vortx.com.br", Name: "Filipe Goeking Costa", "Job Title": "Analista de Segurança", "Manager Name": "Giovanni Gomes Miron", "Manager Email": "ggm@vortx.com.br", Organization: "Vortx", Department: "SEGURANÇA DA INFORMAÇÃO" },
  { Email: "rty@vortx.com.br", Name: "Renato Tyszler", "Job Title": "Diretor Financeiro", "Manager Name": "", "Manager Email": "", Organization: "Vortx", Department: "DIRETORIA" },
  { Email: "jap@vortx.com.br", Name: "José Aparecido da Silva", "Job Title": "Head de Contabilidade", "Manager Name": "Renato Tyszler", "Manager Email": "rty@vortx.com.br", Organization: "Vortx", Department: "CORPORATIVO" },
  { Email: "trm@vortx.com.br", Name: "Thalita Rodrigues De Lira Morais", "Job Title": "Team Leader de Contabilidade", "Manager Name": "José Aparecido da Silva", "Manager Email": "jap@vortx.com.br", Organization: "Vortx", Department: "CONTABILIDADE" }
];

export default function App() {
  const [activeTab, setActiveTab] = useState('upload');
  const [data, setData] = useState([]);
  const [isXlsxLoaded, setIsXlsxLoaded] = useState(false);
  const [error, setError] = useState(null);

  // Load SheetJS dynamically for Excel parsing
  useEffect(() => {
    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    script.async = true;
    script.onload = () => setIsXlsxLoaded(true);
    document.body.appendChild(script);

    return () => {
      document.body.removeChild(script);
    };
  }, []);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!isXlsxLoaded) {
      setError("Biblioteca de processamento carregando. Tente novamente em instantes.");
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = window.XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const parsedData = window.XLSX.utils.sheet_to_json(ws);
        
        // Validate required columns
        if (parsedData.length > 0 && !parsedData[0].hasOwnProperty('Email') && !parsedData[0].hasOwnProperty('Name')) {
          setError("A planilha não possui as colunas necessárias. Baixe o modelo para verificar o formato.");
          return;
        }

        setData(parsedData);
        setError(null);
        setActiveTab('cards');
      } catch (err) {
        setError("Erro ao ler o arquivo. Certifique-se de que é um Excel ou CSV válido.");
      }
    };
    reader.readAsBinaryString(file);
  };

  const loadDemoData = () => {
    setData(MOCK_DATA);
    setActiveTab('cards');
    setError(null);
  };

  const downloadTemplate = () => {
    const csvContent = "data:text/csv;charset=utf-8,Email,Name,Job Title,Manager Name,Manager Email,Organization,Department\negr@vortx.com.br,Everson Gonçalves Ramos,CTO,,,Vortx,PRODUTO & TECNOLOGIA\ndap@vortx.com.br,Diego Pedrosa,Head de Infra.,Everson Gonçalves Ramos,egr@vortx.com.br,Vortx,PRODUTO & TECNOLOGIA";
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "modelo_organoflow.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-800">
      {/* Sidebar / Topbar */}
      <nav className="bg-white border-b border-slate-200 px-6 py-4 flex items-center justify-between sticky top-0 z-10">
        <div className="flex items-center gap-2 text-indigo-600">
          <Network size={28} strokeWidth={2.5} />
          <span className="text-xl font-bold tracking-tight">OrganoFlow</span>
        </div>
        <div className="flex bg-slate-100 p-1 rounded-lg">
          <TabButton 
            active={activeTab === 'upload'} 
            onClick={() => setActiveTab('upload')} 
            icon={<UploadCloud size={18} />} 
            label="Importar" 
          />
          <TabButton 
            active={activeTab === 'cards'} 
            onClick={() => setActiveTab('cards')} 
            icon={<LayoutGrid size={18} />} 
            label="Cartões" 
            disabled={data.length === 0}
          />
          <TabButton 
            active={activeTab === 'hierarchy'} 
            onClick={() => setActiveTab('hierarchy')} 
            icon={<Network size={18} />} 
            label="Hierarquia" 
            disabled={data.length === 0}
          />
        </div>
      </nav>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto p-6">
        {activeTab === 'upload' && (
          <UploadView 
            onUpload={handleFileUpload} 
            onLoadDemo={loadDemoData} 
            onDownloadTemplate={downloadTemplate}
            error={error}
          />
        )}
        
        {activeTab === 'cards' && (
          <CardsView data={data} />
        )}

        {activeTab === 'hierarchy' && (
          <HierarchyView data={data} />
        )}
      </main>
    </div>
  );
}

// --- VIEWS ---

function UploadView({ onUpload, onLoadDemo, onDownloadTemplate, error }) {
  return (
    <div className="max-w-2xl mx-auto mt-12">
      <div className="bg-white p-8 rounded-2xl shadow-sm border border-slate-200 text-center">
        <div className="w-20 h-20 bg-indigo-50 text-indigo-600 rounded-full flex items-center justify-center mx-auto mb-6">
          <UploadCloud size={40} />
        </div>
        <h2 className="text-2xl font-bold mb-2">Importe sua estrutura organizacional</h2>
        <p className="text-slate-500 mb-8">
          Faça upload de uma planilha Excel (.xlsx) ou CSV contendo os colaboradores, seus cargos, departamentos e a quem eles se reportam.
        </p>

        {error && (
          <div className="mb-6 p-4 bg-red-50 text-red-600 rounded-lg text-sm font-medium">
            {error}
          </div>
        )}

        <div className="flex flex-col gap-4 items-center">
          <label className="relative cursor-pointer bg-indigo-600 hover:bg-indigo-700 text-white px-8 py-3 rounded-lg font-medium transition shadow-sm hover:shadow flex items-center gap-2">
            <UploadCloud size={20} />
            Selecionar Arquivo
            <input 
              type="file" 
              accept=".xlsx, .xls, .csv" 
              className="hidden" 
              onChange={onUpload} 
            />
          </label>
          
          <div className="flex items-center gap-4 text-sm mt-4">
            <button 
              onClick={onLoadDemo}
              className="flex items-center gap-2 text-indigo-600 hover:text-indigo-800 font-medium bg-indigo-50 px-4 py-2 rounded-md transition"
            >
              <Play size={16} />
              Testar com Dados de Exemplo
            </button>
            <span className="text-slate-300">|</span>
            <button 
              onClick={onDownloadTemplate}
              className="flex items-center gap-2 text-slate-600 hover:text-slate-900 font-medium transition"
            >
              <Download size={16} />
              Baixar Planilha Modelo
            </button>
          </div>
        </div>
      </div>

      <div className="mt-8 bg-blue-50/50 border border-blue-100 rounded-xl p-6">
        <h3 className="font-semibold text-slate-800 mb-3 flex items-center gap-2">
          <Building size={18} className="text-blue-600" />
          Como estruturar sua planilha:
        </h3>
        <ul className="text-sm text-slate-600 space-y-2 list-disc list-inside">
          <li><strong>Email:</strong> E-mail do colaborador (usado como identificador único).</li>
          <li><strong>Name:</strong> Nome completo do colaborador.</li>
          <li><strong>Job Title:</strong> Título do cargo (ex: Head de Contabilidade).</li>
          <li><strong>Manager Name:</strong> Nome do gerente direto.</li>
          <li><strong>Manager Email:</strong> E-mail do gestor ao qual este colaborador responde. Deixe vazio para líderes sem chefe direto.</li>
          <li><strong>Organization:</strong> Empresa ou organização a que pertence (ex: Vortx).</li>
          <li><strong>Department:</strong> Área de atuação (ex: SEGURANÇA DA INFORMAÇÃO).</li>
        </ul>
      </div>
    </div>
  );
}

function CardsView({ data }) {
  // Process data to group by department and find leaders
  const departmentsInfo = useMemo(() => {
    const depts = {};
    
    // Group by department
    data.forEach(emp => {
      const dName = emp.Department || 'Sem Departamento';
      if (!depts[dName]) depts[dName] = [];
      depts[dName].push(emp);
    });

    // Find the leader for each department
    // A leader is someone who reports to someone OUTSIDE the department, or reports to no one.
    const result = [];
    Object.keys(depts).forEach(deptName => {
      const employees = depts[deptName];
      const deptIds = new Set(employees.map(e => e.Email));
      
      const leaders = employees.filter(e => !e["Manager Email"] || !deptIds.has(e["Manager Email"]));
      // Fallback: if somehow circular or everyone reports inside, pick the first one
      const actualLeaders = leaders.length > 0 ? leaders : [employees[0]];
      
      const members = employees.filter(e => !actualLeaders.includes(e));

      result.push({
        name: deptName,
        leaders: actualLeaders,
        members: members,
        totalCount: employees.length
      });
    });

    // Sort by total count descending
    return result.sort((a, b) => b.totalCount - a.totalCount);
  }, [data]);

  return (
    <div className="animate-in fade-in slide-in-from-bottom-4 duration-500">
      <div className="mb-8">
        <h2 className="text-2xl font-bold text-slate-900">Visão por Departamentos</h2>
        <p className="text-slate-500">Resumo estrutural das áreas da empresa e suas respectivas lideranças.</p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {departmentsInfo.map((dept, idx) => (
          <div key={idx} className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden hover:shadow-md transition">
            {/* Header */}
            <div className="bg-slate-50 border-b border-slate-100 p-5">
              <div className="flex justify-between items-start mb-2">
                <h3 className="text-lg font-bold text-slate-800">{dept.name}</h3>
                <span className="bg-indigo-100 text-indigo-700 text-xs font-bold px-2 py-1 rounded-full flex items-center gap-1">
                  <Users size={12} /> {dept.totalCount}
                </span>
              </div>
            </div>

            {/* Content */}
            <div className="p-5">
              {/* Leaders */}
              <div className="mb-4">
                <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-3">Liderança da Área</p>
                {dept.leaders.map(leader => (
                  <div key={leader.Email} className="flex items-center gap-3 mb-3 last:mb-0">
                    <div className="w-10 h-10 rounded-full bg-gradient-to-tr from-indigo-500 to-purple-500 text-white flex items-center justify-center font-bold text-sm shadow-inner">
                      {getInitials(leader.Name)}
                    </div>
                    <div>
                      <p className="font-semibold text-slate-800 text-sm">{leader.Name}</p>
                      <p className="text-xs text-indigo-600 font-medium">{leader["Job Title"]}</p>
                    </div>
                  </div>
                ))}
              </div>

              {/* Members Preview */}
              {dept.members.length > 0 && (
                <div className="pt-4 border-t border-slate-100">
                  <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-2">Equipe</p>
                  <div className="space-y-2">
                    {dept.members.slice(0, 3).map(member => (
                      <div key={member.Email} className="flex justify-between items-center text-sm">
                        <span className="text-slate-600 truncate mr-2">{member.Name}</span>
                        <span className="text-slate-400 text-xs truncate max-w-[50%] text-right">{member["Job Title"]}</span>
                      </div>
                    ))}
                    {dept.members.length > 3 && (
                      <p className="text-xs text-slate-400 italic pt-1">
                        + {dept.members.length - 3} outros membros...
                      </p>
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function HierarchyView({ data }) {
  const [zoom, setZoom] = useState(1);

  // Build a tree from flat data
  const treeData = useMemo(() => {
    // Deep clone to avoid mutating original state
    const dataCopy = JSON.parse(JSON.stringify(data));
    const idMap = {};
    const roots = [];

    // Map all items by ID
    dataCopy.forEach(item => {
      idMap[item.Email] = item;
      item.children = [];
    });

    // Build hierarchy
    dataCopy.forEach(item => {
      // Treat empty string, "null", or actual null as no manager
      if (!item["Manager Email"] || String(item["Manager Email"]).trim() === '' || item["Manager Email"] === 'null') {
        roots.push(item);
      } else {
        const manager = idMap[item["Manager Email"]];
        if (manager) {
          manager.children.push(item);
        } else {
          // If manager ID exists but not found in data, treat as root to avoid hiding data
          roots.push(item);
        }
      }
    });

    return roots;
  }, [data]);

  const handleZoomIn = () => setZoom(prev => Math.min(prev + 0.2, 2));
  const handleZoomOut = () => setZoom(prev => Math.max(prev - 0.2, 0.4));

  return (
    <div className="animate-in fade-in slide-in-from-bottom-4 duration-500">
      {/* Estilos CSS puros para desenhar as linhas conectivas da árvore */}
      <style>{`
        .org-tree ul {
          padding-top: 24px;
          position: relative;
          display: flex;
          justify-content: center;
          padding-left: 0;
          margin: 0;
        }
        .org-tree li {
          text-align: center;
          list-style-type: none;
          position: relative;
          padding: 24px 10px 0 10px;
        }
        .org-tree li::before, .org-tree li::after {
          content: '';
          position: absolute;
          top: 0;
          right: 50%;
          border-top: 2px solid #cbd5e1;
          width: 50%;
          height: 24px;
          z-index: 1;
        }
        .org-tree li::after {
          right: auto;
          left: 50%;
          border-left: 2px solid #cbd5e1;
        }
        .org-tree li:only-child::after, .org-tree li:only-child::before {
          display: none;
        }
        .org-tree li:only-child {
          padding-top: 0;
        }
        .org-tree li:first-child::before, .org-tree li:last-child::after {
          border: 0 none;
        }
        .org-tree li:last-child::before {
          border-right: 2px solid #cbd5e1;
          border-radius: 0 6px 0 0;
        }
        .org-tree li:first-child::after {
          border-radius: 6px 0 0 0;
        }
        .org-tree ul::before {
          content: '';
          position: absolute;
          top: 0;
          left: 50%;
          border-left: 2px solid #cbd5e1;
          width: 0;
          height: 24px;
          transform: translateX(-50%);
          z-index: 1;
        }
        .org-tree > ul {
          padding-top: 0;
        }
        .org-tree > ul::before {
          display: none;
        }
      `}</style>

      <div className="mb-6 flex flex-col sm:flex-row justify-between items-start sm:items-end gap-4">
        <div>
          <h2 className="text-2xl font-bold text-slate-900">Organograma</h2>
          <p className="text-slate-500">Visão hierárquica em formato de árvore genealógica.</p>
        </div>
        
        {/* Controles de Zoom */}
        <div className="flex items-center gap-1 bg-white p-1 rounded-lg border border-slate-200 shadow-sm">
          <button 
            onClick={handleZoomOut} 
            className="p-1.5 hover:bg-slate-100 rounded text-slate-600 transition-colors"
            title="Diminuir"
          >
            <Minus size={18} />
          </button>
          <span className="text-sm font-medium w-12 text-center text-slate-700">
            {Math.round(zoom * 100)}%
          </span>
          <button 
            onClick={handleZoomIn} 
            className="p-1.5 hover:bg-slate-100 rounded text-slate-600 transition-colors"
            title="Aumentar"
          >
            <Plus size={18} />
          </button>
        </div>
      </div>

      <div className="bg-slate-50 rounded-xl shadow-inner border border-slate-200 p-8 overflow-auto h-[700px] relative cursor-grab active:cursor-grabbing">
        <div 
          className="min-w-max org-tree flex justify-center transform origin-top transition-transform duration-200" 
          style={{ transform: `scale(${zoom})` }}
        >
          {treeData.length > 0 && (
            <ul>
              {treeData.map(rootNode => (
                <TreeNode key={rootNode.Email} node={rootNode} level={0} />
              ))}
            </ul>
          )}
        </div>
      </div>
    </div>
  );
}

// Componente Recursivo para os Cartões da Árvore
function TreeNode({ node, level }) {
  const [isExpanded, setIsExpanded] = useState(true);
  const hasChildren = node.children && node.children.length > 0;

  return (
    <li>
      {/* Container do Cartão */}
      <div className="relative inline-flex flex-col items-center group">
        <div className={`
          w-56 bg-white border rounded-xl shadow-sm p-5 relative z-10 transition-all duration-300
          ${hasChildren && isExpanded ? 'border-indigo-200 shadow-md' : 'border-slate-200 hover:border-indigo-300'}
        `}>
           {/* Avatar */}
           <div className={`w-14 h-14 rounded-full flex items-center justify-center text-lg font-bold mx-auto mb-3
              ${level === 0 ? 'bg-gradient-to-br from-indigo-600 to-purple-600 text-white shadow-md' : 'bg-indigo-50 text-indigo-700 border border-indigo-100'}`}>
              {getInitials(node.Name)}
           </div>

           {/* Informações */}
           <div className="text-center">
              <h4 className="font-bold text-slate-800 text-sm mb-1.5 leading-tight">{node.Name}</h4>
              <p className="text-xs text-indigo-600 font-medium mb-3 leading-tight min-h-[30px] flex items-center justify-center">
                {node["Job Title"]}
              </p>
              <span className="inline-block px-2.5 py-1 bg-slate-100 text-slate-600 text-[10px] rounded uppercase tracking-wider font-bold w-full truncate">
                {node.Department || "Geral"}
              </span>
           </div>

           {/* Botão de Expandir/Recolher */}
           {hasChildren && (
             <button
               onClick={(e) => { e.stopPropagation(); setIsExpanded(!isExpanded); }}
               className="absolute -bottom-3.5 left-1/2 -translate-x-1/2 w-7 h-7 bg-white border border-slate-200 rounded-full flex items-center justify-center text-slate-500 hover:text-indigo-600 hover:border-indigo-300 shadow-sm transition-all z-20 hover:scale-110"
             >
                {isExpanded ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
             </button>
           )}
        </div>

        {/* Badge mostrando quantos filhos estão ocultos quando recolhido */}
        {hasChildren && !isExpanded && (
           <div className="absolute -bottom-9 left-1/2 -translate-x-1/2 text-xs font-bold bg-indigo-100 text-indigo-700 px-2.5 py-0.5 rounded-full shadow-sm z-20 animate-in fade-in zoom-in duration-300">
              +{node.children.length} subordinados
           </div>
        )}
      </div>

      {/* Recursão dos Filhos */}
      {hasChildren && isExpanded && (
         <ul>
           {node.children.map(child => (
             <TreeNode key={child.Email} node={child} level={level + 1} />
           ))}
         </ul>
      )}
    </li>
  );
}

// --- UTILS ---

function TabButton({ active, icon, label, onClick, disabled }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={`
        flex items-center gap-2 px-4 py-2 rounded-md text-sm font-medium transition-all
        ${disabled ? 'opacity-50 cursor-not-allowed text-slate-400' : ''}
        ${!disabled && active ? 'bg-white text-indigo-600 shadow-sm' : ''}
        ${!disabled && !active ? 'text-slate-600 hover:text-slate-900 hover:bg-slate-200/50' : ''}
      `}
    >
      {icon}
      {label}
    </button>
  );
}

function getInitials(name) {
  if (!name) return '??';
  const parts = name.split(' ');
  if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}