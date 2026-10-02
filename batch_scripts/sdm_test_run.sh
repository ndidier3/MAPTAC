#!/bin/bash

# --- Start of SLURM commands -----------

# Request an hour of runtime:
#SBATCH --time=1:00:00

# Use more memory (4GB):
#SBATCH --mem=4G

# Specify a job name:
#SBATCH -J SDMTestRun

# Specify output and error files:
#SBATCH -o batch_out/SDMTestRun-%j.out
#SBATCH -e batch_out/SDMTestRun-%j.err

#SBATCH --partition=batch
#SBATCH --nodes=1
#SBATCH --ntasks-per-node=1

# ----- End of SLURM commands ----

# Load Conda module if available on the cluster, otherwise adjust accordingly:
# module load anaconda3  # Uncomment if Conda needs to be loaded as a module

# Initialize Conda (this sets up Conda for use in this script):
# Commented out because conda automatically loads for my account
# source ~/miniconda3/etc/profile.d/conda.sh  # Update the path if different

# Activate your Conda environment (replace `my_env` with your environment name):
source ~/miniconda3/etc/profile.d/conda.sh
conda deactivate
conda activate sdm-env

# Change to project root (this file lives in batch_scripts/)
cd "$(dirname "$0")/.."

# Run the Test cohort end to end (raw TestData -> Results/Test)
python App/SDM/Scripts/Test/default_analysis.py

# Deactivate the Conda environment after the job finishes:
conda deactivate
