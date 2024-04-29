import numpy as np

class GS_class:
    def __init__(self, num_buses, V_array, Y_bus_matrix, PG_array, QG_array, PL_array, QL_array,
                 Q_min_array, Q_max_array, slack_bus_num, bus_type_array, max_iterations, tolerance):
        print(f"V, {V_array}")
        self.max_iterations, self.tolerance, self.Y_bus_matrix = max_iterations, tolerance, Y_bus_matrix
        self.V_array, self.num_buses = V_array, num_buses
        self.Q_min_array, self.Q_max_array = Q_min_array, Q_max_array
        self.slack_bus_index, self.bus_type_array = slack_bus_num - 1, bus_type_array
        self.V_array_mag, self.V_array_ang = np.abs(V_array), np.angle(V_array)
        self.V_array_mag_copy = np.copy(np.abs(V_array))
        self.P_calc, self.Q_calc = np.zeros(self.num_buses), np.zeros(self.num_buses)
        self.bus_type_array_copy = np.copy(bus_type_array)
        self.P_array = PG_array - PL_array  # Net Active power in p.u.
        self.Q_array = QG_array - QL_array  # Net Reactive power in p.u.
        self.Q_PV = np.zeros(num_buses)

    def GS_solve(self):
        Iteration = 1
        accuracy = 1
        while accuracy >= self.tolerance and Iteration < self.max_iterations:
            V_array_mag_old = np.copy(np.abs(self.V_array))
            for i in range(self.num_buses):
                if self.bus_type_array[i] == 'Slack':
                    continue

                # Vectorized operation for summing the contributions from all other buses
                V_contribution_sum = (np.dot(self.Y_bus_matrix[i, :], self.V_array) -
                                      self.Y_bus_matrix[i, i] * self.V_array[i])

                if self.bus_type_array[i] == 'PV':
                    self.Q_array[i] = - np.imag(np.conj(self.V_array[i]) *
                                 (V_contribution_sum + self.Y_bus_matrix[i][i] * self.V_array[i]))
                    # Enforce Q_min and Q_max limits for iterations until #6
                    if self.Q_max_array[i] != 0:
                        self.Q_array[i] = np.clip(self.Q_array[i], self.Q_min_array[i], self.Q_max_array[i])
                    self.Q_PV[i] = self.Q_array[i]
                    if self.Y_bus_matrix[i][i]:
                        V_new = (1 / self.Y_bus_matrix[i][i]) * ((self.P_array[i] - 1j * self.Q_array[i]) /
                                                                np.conj(self.V_array[i]) - V_contribution_sum)
                    else:
                        V_new = 0
                    self.V_array[i] = V_array_mag_old[i] * (np.cos(np.angle(V_new)) + 1j * np.sin(np.angle(V_new)))
                else:  # PQ bus voltage update
                    if self.Y_bus_matrix[i][i]:
                        V_new = ((1 / self.Y_bus_matrix[i][i]) * ((self.P_array[i] - 1j * self.Q_array[i])
                                                             / np.conj(self.V_array[i]) - V_contribution_sum))
                    else:
                        V_new = 0
                    self.V_array[i] = V_new
            print(f"Iteration {Iteration} V = {self.V_array}")
            accuracy = np.max(np.abs(np.abs(self.V_array) - V_array_mag_old))
            if accuracy < self.tolerance:

                iteration_result = f"Converged after {Iteration} iterations."
                print(iteration_result)
                return self.V_array, self.Q_PV, iteration_result
            Iteration += 1

        iteration_result = "Maximum iterations reached without convergence."
        print(iteration_result)
        return self.V_array, self.Q_PV, iteration_result

    @staticmethod
    def get_iteration_reached(Iteration):
        return Iteration

class NR_class:
    def __init__(self, num_buses, V_array, Y_bus_matrix, PG_array, QG_array, PL_array, QL_array,
                 Q_min_array, Q_max_array, slack_bus_num, bus_type_array, max_iterations, tolerance):
        self.max_iterations, self.tolerance, self.Y_bus_matrix = max_iterations, tolerance, Y_bus_matrix
        self.V_array, self.num_buses = V_array, num_buses
        self.Q_min_array, self.Q_max_array = Q_min_array, Q_max_array
        self.P_sch = PG_array - PL_array
        self.Q_sch = QG_array - QL_array
        self.slack_bus_index, self.bus_type_array = slack_bus_num - 1, bus_type_array
        self.Y_bus_mag, self.Y_bus_ang = np.abs(self.Y_bus_matrix), np.angle(self.Y_bus_matrix)
        self.V_array_mag, self.V_array_ang = np.abs(V_array), np.angle(V_array)
        self.V_array_mag_copy = np.copy(np.abs(V_array))
        self.P_calc, self.Q_calc = np.zeros(self.num_buses), np.zeros(self.num_buses)
        self.bus_type_array_copy = np.copy(bus_type_array)

    def NR_solve(self):
        Iteration = 1
        accuracy = 1
        while accuracy >= self.tolerance and Iteration < self.max_iterations:
            self.calculate_power()
            mismatch = self.update_voltages()
            accuracy = np.linalg.norm(mismatch)
            if accuracy < self.tolerance:  # Check for convergence
                iteration_result = f"Converged after {Iteration} iterations."
                print(iteration_result)
                if not np.array_equal(self.bus_type_array, self.bus_type_array_copy):
                    # method to return changed PV buses to PV and calculate Q_PV with new V_ang and old V_mag
                    self.recalc_Q_PV_after_PQ_treat()

                return self.V_array, iteration_result
            Iteration += 1

        iteration_result = "Maximum iterations reached without convergence."
        print(iteration_result)
        return self.V_array, iteration_result

    def recalc_Q_PV_after_PQ_treat(self):
        for i in range(self.num_buses):
            # Only proceed if the bus is PV
            if self.bus_type_array_copy[i] == 'PV':
                # return changed PV buses
                self.bus_type_array[i] = 'PV'
                # use original V mag for PV with new V ang
                self.V_array_mag[i] = self.V_array_mag_copy[i]
                self.V_array[i] = self.V_array_mag[i] * np.exp(1j * self.V_array_ang[i])

                Q_calc = 0  # Initialize recalculated Q
                for j in range(self.num_buses):
                    # Summation of Yik * Vj for all j
                    Q_calc -= np.imag(self.V_array[i] * np.conj(self.Y_bus_matrix[i, j] * self.V_array[j]))
                # Check Q limit and assign Q_PV
                self.Q_calc[i] = np.clip(Q_calc, self.Q_min_array[i], self.Q_max_array[i])


    def calculate_power(self):
        for i in range(self.num_buses):
            self.P_calc[i] = 0
            self.Q_calc[i] = 0
            for j in range(self.num_buses):
                self.P_calc[i] += (self.V_array_mag[i] * self.V_array_mag[j] * self.Y_bus_mag[i][j] *
                                   np.cos(self.Y_bus_ang[i][j] + self.V_array_ang[j] - self.V_array_ang[i]))
                self.Q_calc[i] -= (self.V_array_mag[i] * self.V_array_mag[j] * self.Y_bus_mag[i][j] *
                                   np.sin(self.Y_bus_ang[i][j] + self.V_array_ang[j] - self.V_array_ang[i]))

            # Q Limit checking for PV buses
            if self.bus_type_array[i] == 'PV' and self.Q_max_array[i] != 0:
                if self.Q_calc[i] > self.Q_max_array[i] or self.Q_calc[i] < self.Q_min_array[i]:
                    self. Q_calc[i] = np.clip(self.Q_calc[i], self.Q_min_array[i], self.Q_max_array[i])
                    self.bus_type_array[i] = 'PQ'   # Temporarily treat as PQ bus
                    self.V_array_mag[i] = self.V_array_mag_copy[i]

    def update_voltages(self):
        non_slack_indices = np.delete(np.arange(self.num_buses), self.slack_bus_index)
        PQ_buses_indices = np.where(self.bus_type_array == 'PQ')[0]
        # Calculate Jacobian
        J = self.calculate_jacobian()
        # Compute mismatches
        DP = self.P_sch[non_slack_indices] - self.P_calc[non_slack_indices]
        DQ = self.Q_sch[PQ_buses_indices] - self.Q_calc[PQ_buses_indices]

        # Mismatch vector for PQ buses
        mismatch = np.concatenate((DP, DQ))  # Excluding slack bus ## DF
        mismatch = mismatch.reshape(-1, 1)  # Reshape DF to be a two-dimensional array

        # Solve for voltage updates DX
        voltage_updates = np.linalg.solve(J, mismatch)
        voltage_updates = voltage_updates.flatten()  # Flatten to 1D
        # Update voltage angles for non-slack buses
        V_ang_update = voltage_updates[:len(non_slack_indices)]  # Extract delta updates for non-slack buses
        self.V_array_ang[non_slack_indices] += V_ang_update

        # Update voltage magnitudes for PQ buses
        voltage_mag_updates = voltage_updates[len(non_slack_indices):]
        self.V_array_mag[PQ_buses_indices] += voltage_mag_updates
        self.V_array = self.V_array_mag * np.exp(1j * self.V_array_ang)

        return mismatch

    def calculate_jacobian(self):
        # Initialize Jacobian sub-matrices
        J1, J2 = np.zeros((self.num_buses, self.num_buses)), np.zeros((self.num_buses, self.num_buses))
        J3, J4 = np.zeros((self.num_buses, self.num_buses)), np.zeros((self.num_buses, self.num_buses))

        for i in range(self.num_buses):
            for j in range(self.num_buses):
                if i != j:
                    angle_diff = self.V_array_ang[i] - self.V_array_ang[j] - self.Y_bus_ang[i][j]
                    # J1ij
                    J1[i][j] = self.V_array_mag[i] * self.V_array_mag[j] * self.Y_bus_mag[i][j] * np.sin(angle_diff)
                    # J2ij
                    J2[i][j] = self.V_array_mag[i] * self.Y_bus_mag[i][j] * np.cos(angle_diff)
                    # J3ij
                    J3[i][j] = - self.V_array_mag[i] * self.V_array_mag[j] * self.Y_bus_mag[i][j] * np.cos(angle_diff)
                    # J4ij
                    J4[i][j] = self.V_array_mag[i] * self.Y_bus_mag[i][j] * np.sin(angle_diff)

        # Calculate diagonal elements of the Jacobian matrix
        for i in range(self.num_buses):
            # J1ii
            J1[i][i] = -self.V_array_mag[i] * sum(self.V_array_mag[j] * self.Y_bus_mag[i][j] *
                        np.sin(self.V_array_ang[i] - self.V_array_ang[j] - self.Y_bus_ang[i, j])
                                                  for j in range(self.num_buses) if j != i)
            # J3ii
            J3[i][i] = self.V_array_mag[i] * sum(self.V_array_mag[j] * self.Y_bus_mag[i][j] *
                        np.cos(self.V_array_ang[i] - self.V_array_ang[j] - self.Y_bus_ang[i, j])
                                                 for j in range(self.num_buses) if j != i)
            # J2ii
            J2[i][i] = self.V_array_mag[i] * self.Y_bus_mag[i][i] * np.cos(self.Y_bus_ang[i][i]) + \
                       sum(self.V_array_mag[j] * self.Y_bus_mag[i][j] * np.cos(
                           self.V_array_ang[i] - self.V_array_ang[j] -
                           self.Y_bus_ang[i][j]) for j in range(self.num_buses))
            # J4ii
            J4[i][i] = -self.V_array_mag[i] * self.Y_bus_mag[i][i] * np.sin(self.Y_bus_ang[i][i]) + \
                       sum(self.V_array_mag[j] * self.Y_bus_mag[i][j] * np.sin(
                           self.V_array_ang[i] - self.V_array_ang[j] -
                           self.Y_bus_ang[i][j]) for j in range(self.num_buses))

        non_slack_indices = np.delete(np.arange(self.num_buses), self.slack_bus_index)
        PQ_buses_indices = np.where(self.bus_type_array == 'PQ')[0]

        J11 = J1[non_slack_indices][:, non_slack_indices]
        J22 = J2[non_slack_indices][:, PQ_buses_indices]
        J33 = J3[PQ_buses_indices][:, non_slack_indices]
        J44 = J4[PQ_buses_indices][:, PQ_buses_indices]
        print('1', J1)
        print('2', J2)
        print('3', J3)
        print('4', J4)

        J = np.block([[J11, J22], [J33, J44]])
        return J


class NRFD_class:
    def __init__(self, num_buses, V_array, Y_bus_matrix, PG_array, QG_array, PL_array, QL_array,
                 Q_min_array, Q_max_array, slack_bus_num, bus_type_array, max_iterations, tolerance):
        self.max_iterations, self.tolerance, self.Y_bus_matrix = max_iterations, tolerance, Y_bus_matrix
        self.V_array, self.num_buses = V_array, num_buses
        self.Q_min_array, self.Q_max_array = Q_min_array, Q_max_array
        self.P_sch = PG_array - PL_array
        self.Q_sch = QG_array - QL_array
        self.slack_bus_index, self.bus_type_array = slack_bus_num - 1, bus_type_array
        self.V_array_mag, self.V_array_ang = np.abs(V_array), np.angle(V_array)
        self.V_array_mag_copy = np.copy(np.abs(V_array))
        self.P_calc, self.Q_calc = np.zeros(self.num_buses), np.zeros(self.num_buses)
        self.bus_type_array_copy = np.copy(bus_type_array)
        self.S_sch = self.P_sch + 1j * self.Q_sch

    def NRFD_solve(self):
        PQ_buses_indices = np.where(self.bus_type_array == 'PQ')[0]
        non_slack_indices = np.delete(np.arange(self.num_buses), self.slack_bus_index)
        Iteration, accuracy = 1, 1
        while accuracy >= self.tolerance and Iteration < self.max_iterations:
            S_BusPowerVector = self.V_array * np.conj(self.Y_bus_matrix @ self.V_array)
            S_mismatch = S_BusPowerVector - self.S_sch
            Q_calc = np.imag(S_BusPowerVector)

            DP = np.real(S_mismatch[non_slack_indices])
            DQ = np.imag(S_mismatch[PQ_buses_indices])

            # Q Limit checking for PV buses
            for i in range(self.num_buses):
                if self.bus_type_array[i] == 'PV' and self.Q_max_array[i] != 0:
                    if Q_calc[i] > self.Q_max_array[i] or Q_calc[i] < self.Q_min_array[i]:
                        Q_calc[i] = np.clip(Q_calc[i], self.Q_min_array[i], self.Q_max_array[i])
                        self.bus_type_array[i] = 'PQ'
                        self.V_array_mag[i] = self.V_array_mag_copy[i]

            B_prime = -np.imag(self.Y_bus_matrix[np.ix_(non_slack_indices, non_slack_indices)])
            B_double_prime = -np.imag(self.Y_bus_matrix[np.ix_(PQ_buses_indices, PQ_buses_indices)])

            # invert of B_prime and B_double_prime:
            inv_B_prime = np.linalg.inv(B_prime)
            inv_B_double_prime = np.linalg.inv(B_double_prime)

            V_ang_updates = -inv_B_prime @ DP
            V_mag_updates = -inv_B_double_prime @ DQ
            self.V_array_ang[non_slack_indices] += V_ang_updates
            self.V_array_mag[PQ_buses_indices] += V_mag_updates
            self.V_array = self.V_array_mag * np.exp(1j * self.V_array_ang)

            Iteration += 1
            accuracy = max(np.max(np.abs(DP)), np.max(np.abs(DQ)))
            if accuracy < self.tolerance:  # Check for convergence
                iteration_result = f"Converged after {Iteration} iterations."

                if not np.array_equal(self.bus_type_array, self.bus_type_array_copy):
                    # method to return changed PV buses to PV and calculate Q_PV with new V_ang and old V_mag
                    self.recalc_Q_PV_after_PQ_treat()
                return self.V_array, iteration_result

            Iteration += 1
        iteration_result = "Maximum iterations reached without convergence."
        print(iteration_result)

        return self.V_array, iteration_result

    def recalc_Q_PV_after_PQ_treat(self):
        for i in range(self.num_buses):
            # Only proceed if the bus is PV
            if self.bus_type_array_copy[i] == 'PV':
                # return changed PV buses
                self.bus_type_array[i] = 'PV'
                # use original V mag for PV with new V ang
                self.V_array_mag[i] = self.V_array_mag_copy[i]
                self.V_array[i] = self.V_array_mag[i] * np.exp(1j * self.V_array_ang[i])

                Q_calc = 0  # Initialize recalculated Q
                for j in range(self.num_buses):
                    # Summation of Yik * Vj for all j
                    Q_calc -= np.imag(self.V_array[i] * np.conj(self.Y_bus_matrix[i, j] * self.V_array[j]))
                # Check Q limit and assign Q_PV
                self.Q_calc[i] = np.clip(Q_calc, self.Q_min_array[i], self.Q_max_array[i])
