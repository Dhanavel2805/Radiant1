pipeline {
    agent any

    triggers {
        githubPush()
    }

    stages {
        stage('Checkout Code') {
            steps {
                checkout scm
            }
        }

        stage('Build') {
            steps {
                echo "Build started..."
                // Example command
                // sh 'mvn clean install'
                // bat 'mvn clean install'
            }
        }

        stage('Test') {
            steps {
                echo "Running tests..."
            }
        }

        stage('Deploy') {
            steps {
                echo "Deployment step..."
            }
        }
    }
}
