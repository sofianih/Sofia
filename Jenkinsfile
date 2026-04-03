// Declarative pipeline: Playwright Python (GC_phs004225.py).
// Prerequisites: Python 3.10+, network to Data Commons, Excel at GC_PHS004225_EXCEL.

pipeline {
    agent any

    environment {
        CI = 'true'
        DATA_COMMONS_URL = 'https://general-qa.datacommons.cancer.gov/#/data'
        // Optional: workbook path on the agent, e.g. "${WORKSPACE}\\testdata\\Base counts for studies.xlsx"
        // GC_PHS004225_EXCEL = ''
    }

    stages {
        stage('Checkout') {
            steps {
                checkout scm
            }
        }

        stage('Python + Playwright') {
            steps {
                script {
                    if (isUnix()) {
                        sh '''
                            python3 -m venv .venv
                            . .venv/bin/activate
                            pip install --upgrade pip
                            pip install -r requirements.txt
                            playwright install chromium
                        '''
                    } else {
                        bat '''
                            python -m venv .venv
                            call .venv\\Scripts\\activate.bat
                            python -m pip install --upgrade pip
                            pip install -r requirements.txt
                            playwright install chromium
                        '''
                    }
                }
            }
        }

        stage('Run phs004225 checks') {
            steps {
                script {
                    if (isUnix()) {
                        sh '''
                            . .venv/bin/activate
                            python GC_phs004225.py
                        '''
                    } else {
                        bat '''
                            call .venv\\Scripts\\activate.bat
                            python GC_phs004225.py
                        '''
                    }
                }
            }
        }
    }

    post {
        always {
            archiveArtifacts artifacts: 'QA_Report.html', allowEmptyArchive: true
        }
    }
}
